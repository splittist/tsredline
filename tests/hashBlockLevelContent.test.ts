import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';

import { hashBlockLevelContent } from '../src';
import { loadFixture } from './helpers/fixtureLoader';

async function createSyntheticDocx(documentXml: string): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('hashBlockLevelContent', () => {
  it('generates deterministic block hashes for the same DOCX', async () => {
    const docx = await loadFixture('WC', 'WC002-Unmodified.docx');

    const first = await hashBlockLevelContent(docx, {
      ignoreWhitespace: true,
    });
    const second = await hashBlockLevelContent(docx, {
      ignoreWhitespace: true,
    });

    expect(first.blockCount).toBeGreaterThan(0);
    expect(first.blocks.map((b) => b.hash)).toEqual(second.blocks.map((b) => b.hash));
  });

  it('produces different hash signatures for modified fixtures', async () => {
    const baseline = await loadFixture('WC', 'WC002-Unmodified.docx');
    const candidate = await loadFixture('WC', 'WC002-DiffInMiddle.docx');

    const baselineHashes = await hashBlockLevelContent(baseline, {
      ignoreWhitespace: true,
    });
    const candidateHashes = await hashBlockLevelContent(candidate, {
      ignoreWhitespace: true,
    });

    expect(baselineHashes.blocks.map((b) => b.hash)).not.toEqual(
      candidateHashes.blocks.map((b) => b.hash),
    );
  });

  it('can include table-row hashes for table-heavy fixtures', async () => {
    const tableDocx = await loadFixture('WC', 'WC024-Table-Before.docx');

    const result = await hashBlockLevelContent(tableDocx, {
      includeTableRows: true,
      ignoreWhitespace: true,
    });

    const rowCount = result.blocks.filter((b) => b.kind === 'table-row').length;
    expect(result.blockCount).toBeGreaterThan(0);
    expect(rowCount).toBeGreaterThan(0);
  });

  it('uses accept mode by default for revision-marked content', async () => {
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

    const result = await hashBlockLevelContent(docx, {
      ignoreWhitespace: true,
    });

    expect(result.blocks[0]?.text).toBe('Alpha Beta Delta');
  });

  it('changes block signatures between accept and reject revision modes', async () => {
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

    const accepted = await hashBlockLevelContent(docx, {
      ignoreWhitespace: true,
      revisionMode: 'accept',
    });
    const rejected = await hashBlockLevelContent(docx, {
      ignoreWhitespace: true,
      revisionMode: 'reject',
    });

    expect(accepted.blocks[0]?.text).toBe('Alpha Beta Delta');
    expect(rejected.blocks[0]?.text).toBe('Alpha Gamma Delta');
    expect(accepted.blocks[0]?.hash).not.toBe(rejected.blocks[0]?.hash);
  });
});
