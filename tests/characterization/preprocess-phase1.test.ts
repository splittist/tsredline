import { describe, expect, it } from 'vitest';

import { compareDocx } from '../../src';
import { loadFixture } from '../helpers/fixtureLoader';

describe('phase-1 preprocess comparison', () => {
  it('uses preprocessed-text mode on real DOCX fixtures', async () => {
    const baseline = await loadFixture('WC', 'WC002-Unmodified.docx');
    const candidate = await loadFixture('WC', 'WC002-Unmodified.docx');

    const result = await compareDocx(baseline, candidate, {
      includeDiagnostics: true,
      ignoreWhitespace: true,
    });

    expect(result.equal).toBe(true);
    expect(result.metadata.comparisonMode).toBe('preprocessed-text');
    expect(result.metadata.baselineParagraphs).toBeGreaterThan(0);
    expect(result.metadata.candidateParagraphs).toBeGreaterThan(0);
  });

  it('reports semantic text replacement for differing fixtures', async () => {
    const baseline = await loadFixture('CA', 'CA001-Plain.docx');
    const candidate = await loadFixture('CA', 'CA001-Plain-Mod.docx');

    const result = await compareDocx(baseline, candidate, {
      includeDiagnostics: true,
      ignoreWhitespace: true,
    });

    expect(result.equal).toBe(false);
    expect(result.metadata.comparisonMode).toBe('preprocessed-text');
    expect(result.changes[0]?.path).toBe('word/document.xml:text');
    expect(result.changes[0]?.kind).toBe('replace');
  });
});
