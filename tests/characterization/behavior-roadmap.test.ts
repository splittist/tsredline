import { describe, expect, test } from 'vitest';

import { compareDocx } from '../../src';
import { CHARACTERIZATION_CASES } from '../fixtures/caseManifest';
import { loadFixture } from '../helpers/fixtureLoader';

describe('characterization behavior roadmap', () => {
  test('has a curated C#-traceable fixture matrix', () => {
    expect(CHARACTERIZATION_CASES.length).toBeGreaterThanOrEqual(10);
  });

  for (const fixtureCase of CHARACTERIZATION_CASES) {
    test(
      `${fixtureCase.id} (${fixtureCase.category}): ${fixtureCase.expectedBehavior}`,
      async () => {
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

        // These are current-engine invariants, not final C# parity assertions.
        expect(result.metadata.comparisonMode).toBe('preprocessed-text');

        expect(result.metadata.baselineUnits).toBeDefined();
        expect(result.metadata.candidateUnits).toBeDefined();
        expect(result.metadata.equalUnits).toBeDefined();
        expect(result.metadata.deletedUnits).toBeDefined();
        expect(result.metadata.insertedUnits).toBeDefined();

        expect(result.metadata.baselineComparisonUnits).toBeDefined();
        expect(result.metadata.candidateComparisonUnits).toBeDefined();
        expect(result.metadata.equalComparisonUnits).toBeDefined();
        expect(result.metadata.deletedComparisonUnits).toBeDefined();
        expect(result.metadata.insertedComparisonUnits).toBeDefined();

        expect(result.metadata.equalUnits! + result.metadata.deletedUnits!).toBe(
          result.metadata.baselineUnits,
        );
        expect(result.metadata.equalUnits! + result.metadata.insertedUnits!).toBe(
          result.metadata.candidateUnits,
        );

        expect(
          result.metadata.equalComparisonUnits! +
            result.metadata.deletedComparisonUnits!,
        ).toBe(result.metadata.baselineComparisonUnits);
        expect(
          result.metadata.equalComparisonUnits! +
            result.metadata.insertedComparisonUnits!,
        ).toBe(result.metadata.candidateComparisonUnits);

        expect(result.notices.some((n) => n.code === 'SKELETON_ENGINE')).toBe(true);
      },
    );
  }
});
