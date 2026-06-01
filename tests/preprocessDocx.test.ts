import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';

import { preprocessDocx } from '../src';
import { loadFixture } from './helpers/fixtureLoader';

async function createSyntheticDocx(documentXml: string): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('preprocessDocx', () => {
  it('extracts paragraph and word counts from a DOCX fixture', async () => {
    const docx = await loadFixture('CA', 'CA001-Plain.docx');

    const result = await preprocessDocx(docx, {
      ignoreWhitespace: true,
    });

    expect(result.paragraphCount).toBeGreaterThan(0);
    expect(result.wordCount).toBeGreaterThan(0);
    expect(result.normalizedText.length).toBeGreaterThan(0);
  });

  it('accepts inserted text and drops deleted text by default', async () => {
    const docx = await createSyntheticDocx(`
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r><w:t>Alpha </w:t></w:r>
            <w:ins><w:r><w:t>Beta </w:t></w:r></w:ins>
            <w:del><w:r><w:delText>Gamma </w:delText></w:r></w:del>
            <w:r><w:t>Delta</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    `);

    const result = await preprocessDocx(docx, {
      ignoreWhitespace: true,
    });

    expect(result.normalizedText).toBe('Alpha Beta Delta');
  });

  it('rejects inserted text and restores deleted text in reject mode', async () => {
    const docx = await createSyntheticDocx(`
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r><w:t>Alpha </w:t></w:r>
            <w:ins><w:r><w:t>Beta </w:t></w:r></w:ins>
            <w:del><w:r><w:delText>Gamma </w:delText></w:r></w:del>
            <w:r><w:t>Delta</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    `);

    const result = await preprocessDocx(docx, {
      ignoreWhitespace: true,
      revisionMode: 'reject',
    });

    expect(result.normalizedText).toBe('Alpha Gamma Delta');
  });
});
