import JSZip from 'jszip';

import { collectContentParts } from './docxParts';
import { preprocessRevisionMarkup, type RevisionMode } from './preprocessDocx';
import { WCT_NS, WCT_UNID_LOCAL } from './assignUnids';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

export type AtomKind =
  | 'word'
  | 'image'
  | 'math'
  | 'field-char'
  | 'break'
  | 'symbol'
  | 'other-inline';

export interface ComparisonUnitAtom {
  readonly kind: AtomKind;
  // For 'word': the word string. For non-text kinds: serialized element XML
  // (used as the identity key during LCS comparison).
  readonly text: string;
  // wct:id of the enclosing w:p, as stamped by assignUnids. 0 if unstamped.
  readonly paraUnid: number;
  // Serialized w:rPr of the containing run, or '' if absent / paragraph-level atom.
  readonly runPropsXml: string;
  // Serialized w:pPr of the containing paragraph, or ''.
  readonly paraPropsXml: string;
  readonly partName: string;
  // Ancestor grouping keys from outermost to innermost, e.g. ["tbl:5","tr:8","tc:11","p:14"].
  // Used by getComparisonUnitList to build the paragraph/cell/row/table hierarchy.
  readonly ancestorKeys: readonly string[];
}

export interface CreateAtomListOptions {
  readonly revisionMode?: RevisionMode;
}

export interface AtomListResult {
  readonly atoms: readonly ComparisonUnitAtom[];
  readonly atomCount: number;
  readonly partNames: readonly string[];
}

// ---------------------------------------------------------------------------
// DOM helpers
// ---------------------------------------------------------------------------

function localName(el: Element): string {
  const n = el.nodeName;
  const c = n.indexOf(':');
  return c >= 0 ? n.slice(c + 1) : n;
}

function serializeEl(el: Element): string {
  return new XMLSerializer().serializeToString(el);
}

function firstChildXml(parent: Element, childLocal: string): string {
  for (const child of parent.childNodes) {
    if (child.nodeType === Node.ELEMENT_NODE && localName(child as Element) === childLocal) {
      return serializeEl(child as Element);
    }
  }
  return '';
}

function paraUnidOf(el: Element): number {
  const raw = el.getAttributeNS(WCT_NS, WCT_UNID_LOCAL);
  if (raw === null) return 0;
  const n = parseInt(raw, 10);
  return Number.isFinite(n) && n > 0 ? n : 0;
}

// Elements that form the structural grouping hierarchy for comparison units.
// Matches C# ComparisonGroupingElements: W.p, W.tbl, W.tr, W.tc, W.txbxContent.
const GROUPING_LOCALS = new Set(['p', 'tbl', 'tr', 'tc', 'txbxContent']);

// Returns ancestor grouping keys for a paragraph element, outermost first.
// E.g. ["tbl:5","tr:8","tc:11","p:14"] for a paragraph inside a table cell.
function ancestorKeysOf(para: Element): readonly string[] {
  const chain: string[] = [];
  let el: Element | null = para;
  while (el !== null) {
    const local = localName(el);
    if (local === 'body' || local === 'footnotes' || local === 'endnotes' || local === 'document') break;
    if (GROUPING_LOCALS.has(local)) {
      const unid = el.getAttributeNS(WCT_NS, WCT_UNID_LOCAL) ?? '0';
      chain.push(`${local}:${unid}`);
    }
    el = el.parentElement;
  }
  chain.reverse();
  return chain;
}

// ---------------------------------------------------------------------------
// Atom-kind tables
// ---------------------------------------------------------------------------

// Run child elements that produce a single non-text atom.
const RUN_INLINE_KINDS: Readonly<Record<string, AtomKind>> = {
  drawing: 'image',
  pict: 'image',
  sym: 'symbol',
  br: 'break',
  fldChar: 'field-char',
};

// Run child elements that are ignored (markup/hints carry no comparison value).
const SKIP_IN_RUN = new Set([
  'rPr',
  'instrText',
  'tab',
  'noBreakHyphen',
  'softHyphen',
  'lastRenderedPageBreak',
  'cr',
  'delText',
]);

// Paragraph child elements that are transparent containers of runs.
const RUN_CONTAINERS = new Set([
  'hyperlink',
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'sdt',
  'sdtContent',
  'fldSimple',
  'customXml',
]);

// Paragraph child elements that carry no content atoms.
const SKIP_IN_PARA = new Set([
  'pPr',
  'bookmarkStart',
  'bookmarkEnd',
  'proofErr',
  'permStart',
  'permEnd',
]);

// ---------------------------------------------------------------------------
// Atom extraction
// ---------------------------------------------------------------------------

function atomsFromRun(
  run: Element,
  paraUnid: number,
  paraPropsXml: string,
  ancestorKeys: readonly string[],
  partName: string,
): ComparisonUnitAtom[] {
  const runPropsXml = firstChildXml(run, 'rPr');
  const atoms: ComparisonUnitAtom[] = [];

  for (const child of run.childNodes) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const el = child as Element;
    const local = localName(el);

    if (SKIP_IN_RUN.has(local)) continue;

    if (local === 't') {
      const words = (el.textContent ?? '').split(/\s+/).filter((w) => w.length > 0);
      for (const word of words) {
        atoms.push({ kind: 'word', text: word, paraUnid, runPropsXml, paraPropsXml, partName, ancestorKeys });
      }
      continue;
    }

    const inlineKind = RUN_INLINE_KINDS[local];
    if (inlineKind !== undefined) {
      // Non-text inlines are identified by their serialized XML; they carry
      // the run's rPr so reconstruction can wrap them in the correct w:r.
      atoms.push({ kind: inlineKind, text: serializeEl(el), paraUnid, runPropsXml, paraPropsXml, partName, ancestorKeys });
      continue;
    }

    // Unrecognised run child — emit as other-inline so it round-trips.
    atoms.push({ kind: 'other-inline', text: serializeEl(el), paraUnid, runPropsXml, paraPropsXml, partName, ancestorKeys });
  }

  return atoms;
}

function collectAtomsFromContainer(
  container: Element,
  paraUnid: number,
  paraPropsXml: string,
  ancestorKeys: readonly string[],
  partName: string,
  out: ComparisonUnitAtom[],
): void {
  for (const child of container.childNodes) {
    if (child.nodeType !== Node.ELEMENT_NODE) continue;
    const el = child as Element;
    const local = localName(el);

    if (SKIP_IN_PARA.has(local)) continue;

    if (local === 'r') {
      out.push(...atomsFromRun(el, paraUnid, paraPropsXml, ancestorKeys, partName));
      continue;
    }

    if (RUN_CONTAINERS.has(local)) {
      collectAtomsFromContainer(el, paraUnid, paraPropsXml, ancestorKeys, partName, out);
      continue;
    }

    // m:oMath is a paragraph-level math island — not inside a w:r.
    if (el.namespaceURI === M_NS && local === 'oMath') {
      out.push({ kind: 'math', text: serializeEl(el), paraUnid, runPropsXml: '', paraPropsXml, partName, ancestorKeys });
      continue;
    }
    // Anything else at paragraph level is silently skipped.
  }
}

function atomsFromParagraph(para: Element, partName: string): ComparisonUnitAtom[] {
  const paraUnid = paraUnidOf(para);
  const paraPropsXml = firstChildXml(para, 'pPr');
  const ancestorKeys = ancestorKeysOf(para);
  const atoms: ComparisonUnitAtom[] = [];
  collectAtomsFromContainer(para, paraUnid, paraPropsXml, ancestorKeys, partName, atoms);
  return atoms;
}

function atomsFromXml(xml: string, partName: string): ComparisonUnitAtom[] {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`XML parse error in content part: ${partName}`);
  }

  const atoms: ComparisonUnitAtom[] = [];
  for (const para of doc.getElementsByTagNameNS(W_NS, 'p')) {
    atoms.push(...atomsFromParagraph(para, partName));
  }
  return atoms;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export async function createAtomList(
  docx: ArrayBuffer,
  options: CreateAtomListOptions = {},
): Promise<AtomListResult> {
  const revisionMode = options.revisionMode ?? 'accept';
  const zip = await JSZip.loadAsync(new Uint8Array(docx));
  const parts = collectContentParts(zip);
  const allAtoms: ComparisonUnitAtom[] = [];
  const partNames: string[] = [];

  for (const partName of parts) {
    const entry = zip.file(partName);
    if (entry === null) continue;

    const rawXml = await entry.async('text');
    const processedXml = preprocessRevisionMarkup(rawXml, revisionMode);
    allAtoms.push(...atomsFromXml(processedXml, partName));
    partNames.push(partName);
  }

  return {
    atoms: allAtoms,
    atomCount: allAtoms.length,
    partNames,
  };
}
