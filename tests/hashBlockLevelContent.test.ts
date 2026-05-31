import { describe, expect, it } from 'vitest';

import { hashBlockLevelContent } from '../src';
import { loadFixture } from './helpers/fixtureLoader';

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
});
