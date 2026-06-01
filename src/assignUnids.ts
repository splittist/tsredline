import JSZip from 'jszip';

export const WCT_NS = 'urn:tsredline:wct';
export const WCT_PREFIX = 'wct';
export const WCT_UNID_LOCAL = 'id';
export const WCT_UNID_QNAME = `${WCT_PREFIX}:${WCT_UNID_LOCAL}`;
export const WCT_XMLNS_QNAME = `xmlns:${WCT_PREFIX}`;

// Block-level local names that receive a stamped unid.
// w:p and w:tr are the primary correlation units for phase-1 LCS matching.
const STAMPED_LOCAL_NAMES = new Set(['p', 'tr']);

// Fixed optional parts, in processing order after word/document.xml.
const FIXED_OPTIONAL_PARTS = [
  'word/footnotes.xml',
  'word/endnotes.xml',
  'word/comments.xml',
] as const;

// Header and footer parts are discovered from the zip manifest.
const HEADER_FOOTER_RE = /^word\/(header|footer)\d+\.xml$/;

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

function collectParts(zip: JSZip): string[] {
  const parts: string[] = ['word/document.xml'];

  for (const partName of FIXED_OPTIONAL_PARTS) {
    if (zip.file(partName) !== null) {
      parts.push(partName);
    }
  }

  for (const partName of Object.keys(zip.files).sort()) {
    if (HEADER_FOOTER_RE.test(partName)) {
      parts.push(partName);
    }
  }

  return parts;
}

export async function assignUnids(docx: ArrayBuffer): Promise<AssignUnidsResult> {
  const zip = await JSZip.loadAsync(new Uint8Array(docx));
  const parts = collectParts(zip);
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
