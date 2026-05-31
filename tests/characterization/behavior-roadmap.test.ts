import { describe, expect, test } from 'vitest';

import { CHARACTERIZATION_CASES } from '../fixtures/caseManifest';

describe('characterization behavior roadmap', () => {
  test('has a curated C#-traceable fixture matrix', () => {
    expect(CHARACTERIZATION_CASES.length).toBeGreaterThanOrEqual(10);
  });

  for (const fixtureCase of CHARACTERIZATION_CASES) {
    test.todo(
      `${fixtureCase.id} (${fixtureCase.category}): ${fixtureCase.expectedBehavior}`,
    );
  }
});
