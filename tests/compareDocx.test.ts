import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';

import { compareDocx } from '../src';

function toArrayBuffer(value: string): ArrayBuffer {
  return new TextEncoder().encode(value).buffer;
}

async function createSyntheticDocx(documentXml: string): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('compareDocx', () => {
  it('reports identical inputs as equal', async () => {
    const result = await compareDocx(
      toArrayBuffer('same document'),
      toArrayBuffer('same document'),
    );

    expect(result.equal).toBe(true);
    expect(result.changes).toHaveLength(0);
    expect(result.metadata.identicalBinary).toBe(true);
    expect(result.notices[0]?.code).toBe('SKELETON_ENGINE');
  });

  it('reports differing inputs as a placeholder replacement', async () => {
    const result = await compareDocx(
      toArrayBuffer('baseline document'),
      toArrayBuffer('candidate document'),
    );

    expect(result.equal).toBe(false);
    expect(result.changes).toEqual([
      {
        kind: 'replace',
        path: 'docx-package',
        before: '17 bytes',
        after: '18 bytes',
      },
    ]);
    expect(result.metadata.identicalBinary).toBe(false);
  });

  it('reports DOCX-mode correlation unit counts for text replacements', async () => {
    const baseline = await createSyntheticDocx(`
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>alpha beta gamma</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const candidate = await createSyntheticDocx(`
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>alpha delta gamma</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);

    const result = await compareDocx(baseline, candidate);

    expect(result.equal).toBe(false);
    expect(result.metadata.comparisonMode).toBe('preprocessed-text');
    expect(result.metadata.baselineUnits).toBe(3);
    expect(result.metadata.candidateUnits).toBe(3);
    expect(result.metadata.equalUnits).toBe(2);
    expect(result.metadata.deletedUnits).toBe(1);
    expect(result.metadata.insertedUnits).toBe(1);
    expect(result.changes).toContainEqual({
      kind: 'delete',
      path: 'word/document.xml:correlation',
      before: '1 preprocessed unit(s)',
    });
    expect(result.changes).toContainEqual({
      kind: 'insert',
      path: 'word/document.xml:correlation',
      after: '1 preprocessed unit(s)',
    });
  });

  it('reports all units as equal for matching DOCX text', async () => {
    const baseline = await createSyntheticDocx(`
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>alpha beta gamma</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);
    const candidate = await createSyntheticDocx(`
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>alpha beta gamma</w:t></w:r></w:p>
        </w:body>
      </w:document>
    `);

    const result = await compareDocx(baseline, candidate);

    expect(result.equal).toBe(true);
    expect(result.metadata.comparisonMode).toBe('preprocessed-text');
    expect(result.metadata.baselineUnits).toBe(3);
    expect(result.metadata.candidateUnits).toBe(3);
    expect(result.metadata.equalUnits).toBe(3);
    expect(result.metadata.deletedUnits).toBe(0);
    expect(result.metadata.insertedUnits).toBe(0);
  });
});
