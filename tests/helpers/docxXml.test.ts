import { describe, expect, it } from 'vitest';

import { loadFixture } from './fixtureLoader';
import {
  countXmlTagInPart,
  hasDocxPart,
  listDocxParts,
  readDocxPartText,
} from './docxXml';

describe('docxXml helper', () => {
  it('lists common parts in a fixture docx', async () => {
    const docx = await loadFixture('CA', 'CA001-Plain.docx');
    const parts = await listDocxParts(docx);

    expect(parts).toContain('[Content_Types].xml');
    expect(parts).toContain('word/document.xml');
  });

  it('reads XML text from word/document.xml', async () => {
    const docx = await loadFixture('WC', 'WC002-Unmodified.docx');
    const documentXml = await readDocxPartText(docx, 'word/document.xml');

    expect(documentXml).toContain('<w:document');
    expect(documentXml).toContain('<w:body');
  });

  it('detects optional parts and counts tags', async () => {
    const docx = await loadFixture('WC', 'WC020-FootNote-Before.docx');
    const hasFootnotes = await hasDocxPart(docx, 'word/footnotes.xml');

    expect(hasFootnotes).toBe(true);

    const paragraphCount = await countXmlTagInPart(
      docx,
      'word/document.xml',
      'w:p',
    );

    expect(paragraphCount).toBeGreaterThan(0);
  });
});
