import { describe, expect, it } from 'vitest';

import { preprocessDocx } from '../src';
import { loadFixture } from './helpers/fixtureLoader';

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
});
