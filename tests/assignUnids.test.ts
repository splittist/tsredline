import { describe, expect, it } from 'vitest';

import { WCT_NS, WCT_UNID_LOCAL, assignUnids, readUnid } from '../src/assignUnids';
import { loadFixture } from './helpers/fixtureLoader';
import { readDocxPartText } from './helpers/docxXml';

function parseXml(xml: string): Document {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error('parsererror in test XML');
  }
  return doc;
}

function collectUnids(doc: Document, localName: string): number[] {
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const elements = Array.from(doc.getElementsByTagNameNS(ns, localName));
  return elements
    .map((el) => readUnid(el))
    .filter((id): id is number => id !== null);
}

describe('assignUnids', () => {
  it('stamps every w:p in document.xml with a unique positive integer', async () => {
    const docx = await loadFixture('WC', 'WC002-Unmodified.docx');
    const result = await assignUnids(docx);

    const xml = await readDocxPartText(result.docx, 'word/document.xml');
    const doc = parseXml(xml);
    const ids = collectUnids(doc, 'p');

    expect(ids.length).toBeGreaterThan(0);
    expect(new Set(ids).size).toBe(ids.length);
    expect(Math.min(...ids)).toBe(1);
  });

  it('stamps w:tr elements in a table fixture', async () => {
    const docx = await loadFixture('WC', 'WC024-Table-Before.docx');
    const result = await assignUnids(docx);

    const xml = await readDocxPartText(result.docx, 'word/document.xml');
    const doc = parseXml(xml);
    const rowIds = collectUnids(doc, 'tr');

    expect(rowIds.length).toBeGreaterThan(0);
    expect(new Set(rowIds).size).toBe(rowIds.length);
  });

  it('ids are globally unique across w:p and w:tr within a part', async () => {
    const docx = await loadFixture('WC', 'WC024-Table-Before.docx');
    const result = await assignUnids(docx);

    const xml = await readDocxPartText(result.docx, 'word/document.xml');
    const doc = parseXml(xml);
    const pIds = collectUnids(doc, 'p');
    const trIds = collectUnids(doc, 'tr');
    const all = [...pIds, ...trIds];

    expect(new Set(all).size).toBe(all.length);
  });

  it('processes footnotes.xml when present and continues the counter', async () => {
    const docx = await loadFixture('WC', 'WC020-FootNote-Before.docx');
    const result = await assignUnids(docx);

    expect(result.partsStamped).toContain('word/document.xml');
    expect(result.partsStamped).toContain('word/footnotes.xml');

    const docXml = await readDocxPartText(result.docx, 'word/document.xml');
    const fnXml = await readDocxPartText(result.docx, 'word/footnotes.xml');

    const docIds = collectUnids(parseXml(docXml), 'p');
    const fnIds = collectUnids(parseXml(fnXml), 'p');

    // No id overlap between parts.
    const docSet = new Set(docIds);
    expect(fnIds.every((id) => !docSet.has(id))).toBe(true);

    // Footnote ids are all greater than every document id.
    expect(Math.min(...fnIds)).toBeGreaterThan(Math.max(...docIds));
  });

  it('unidCount equals total elements stamped across all parts', async () => {
    const docx = await loadFixture('WC', 'WC020-FootNote-Before.docx');
    const result = await assignUnids(docx);

    const stampedKinds = ['p', 'tr', 'tbl', 'tc', 'txbxContent'];
    let total = 0;
    for (const partName of result.partsStamped) {
      const xml = await readDocxPartText(result.docx, partName);
      const doc = parseXml(xml);
      for (const kind of stampedKinds) {
        total += collectUnids(doc, kind).length;
      }
    }

    expect(result.unidCount).toBe(total);
  });

  it('returns a valid DOCX that still contains word/document.xml', async () => {
    const docx = await loadFixture('CA', 'CA001-Plain.docx');
    const result = await assignUnids(docx);

    const xml = await readDocxPartText(result.docx, 'word/document.xml');
    expect(xml).toContain('<w:document');
    expect(xml).toContain(WCT_NS);
    expect(xml).toContain(WCT_UNID_LOCAL);
  });

  it('stamping is idempotent: re-stamping overwrites existing ids monotonically', async () => {
    const docx = await loadFixture('WC', 'WC002-Unmodified.docx');
    const first = await assignUnids(docx);
    const second = await assignUnids(first.docx);

    // unidCount should be the same because the same elements get stamped.
    expect(second.unidCount).toBe(first.unidCount);

    const xml = await readDocxPartText(second.docx, 'word/document.xml');
    const ids = collectUnids(parseXml(xml), 'p');
    expect(new Set(ids).size).toBe(ids.length);
    expect(Math.min(...ids)).toBe(1);
  });
});
