import JSZip from 'jszip';

import { collectContentParts } from './docxParts';

export const WCT_NS = 'urn:tsredline:wct';
export const WCT_PREFIX = 'wct';
export const WCT_UNID_LOCAL = 'id';
export const WCT_UNID_QNAME = `${WCT_PREFIX}:${WCT_UNID_LOCAL}`;
export const WCT_XMLNS_QNAME = `xmlns:${WCT_PREFIX}`;

// Block-level local names that receive a stamped unid.
// p/tr are the primary LCS correlation units; tbl/tc/txbxContent are needed
// to build the hierarchical grouping keys used by getComparisonUnitList.
const STAMPED_LOCAL_NAMES = new Set(['p', 'tr', 'tbl', 'tc', 'txbxContent']);

export interface AssignUnidsResult {
  readonly docx: ArrayBuffer;
  readonly partsStamped: readonly string[];
  readonly unidCount: number;
}

function wordLocalName(element: Element): string {
  const name = element.nodeName;
  const colon = name.indexOf(':');
  return colon >= 0 ? name.slice(colon + 1) : name;
}

function stampPart(xml: string, counter: { value: number }): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');

  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error('XML parse error while stamping unids');
  }

  // Declare the wct namespace on the root so serializers emit it once there.
  doc.documentElement.setAttributeNS(
    'http://www.w3.org/2000/xmlns/',
    WCT_XMLNS_QNAME,
    WCT_NS,
  );

  const walker = doc.createTreeWalker(doc.documentElement, NodeFilter.SHOW_ELEMENT);
  let node: Node | null = walker.currentNode;
  while (node !== null) {
    const element = node as Element;
    if (STAMPED_LOCAL_NAMES.has(wordLocalName(element))) {
      element.setAttributeNS(WCT_NS, WCT_UNID_QNAME, String(counter.value));
      counter.value += 1;
    }
    node = walker.nextNode();
  }

  return new XMLSerializer().serializeToString(doc);
}

export async function assignUnids(docx: ArrayBuffer): Promise<AssignUnidsResult> {
  const zip = await JSZip.loadAsync(new Uint8Array(docx));
  const parts = collectContentParts(zip);
  const counter = { value: 1 };
  const partsStamped: string[] = [];

  for (const partName of parts) {
    const entry = zip.file(partName);
    if (entry === null) continue;

    const xml = await entry.async('text');
    const stamped = stampPart(xml, counter);
    zip.file(partName, stamped);
    partsStamped.push(partName);
  }

  return {
    docx: await zip.generateAsync({ type: 'arraybuffer' }),
    partsStamped,
    unidCount: counter.value - 1,
  };
}

export function readUnid(element: Element): number | null {
  const raw = element.getAttributeNS(WCT_NS, WCT_UNID_LOCAL);
  if (raw === null) return null;
  const n = Number(raw);
  return Number.isInteger(n) && n > 0 ? n : null;
}
