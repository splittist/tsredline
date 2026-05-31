import { describe, expect, it } from 'vitest';

import { compareDocx } from '../../src';
import { CHARACTERIZATION_CASES } from '../fixtures/caseManifest';
import { loadFixture } from '../helpers/fixtureLoader';

describe('fixture-based characterization smoke', () => {
  it.each(CHARACTERIZATION_CASES)(
    '$id should compare without throwing on real DOCX fixtures',
    async (fixtureCase) => {
      const baseline = await loadFixture(
        fixtureCase.baseline.group,
        fixtureCase.baseline.file,
      );
      const candidate = await loadFixture(
        fixtureCase.candidate.group,
        fixtureCase.candidate.file,
      );

      const result = await compareDocx(baseline, candidate, {
        includeDiagnostics: true,
        trackMoves: true,
      });

      expect(result.metadata.baselineSize).toBeGreaterThan(0);
      expect(result.metadata.candidateSize).toBeGreaterThan(0);
      expect(result.notices.some((n) => n.code === 'SKELETON_ENGINE')).toBe(true);
    },
  );
});
