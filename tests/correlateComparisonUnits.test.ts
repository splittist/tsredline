import { describe, expect, it } from 'vitest';

import {
  correlateComparisonUnits,
  doLcsAlgorithm,
  doLcsAlgorithmForTable,
  findCommonAtBeginningAndEnd,
  processCorrelatedHashes,
  type CorrelatedSequence,
  type CorrelationStatus,
} from '../src/correlateComparisonUnits';
import type {
  ComparisonUnit,
  ComparisonUnitGroup,
  ComparisonUnitWord,
} from '../src/getComparisonUnitList';

function makeWord(hash: string): ComparisonUnitWord {
  return {
    kind: 'word',
    sha1Hash: hash,
    atoms: [],
  };
}

function statuses(items: readonly { correlationStatus: CorrelationStatus }[]): CorrelationStatus[] {
  return items.map((i) => i.correlationStatus);
}

function unknown(left: readonly string[], right: readonly string[]): CorrelatedSequence {
  return {
    correlationStatus: 'unknown',
    comparisonUnits1: left.map(makeWord),
    comparisonUnits2: right.map(makeWord),
  };
}

function makeRow(wordHashes: readonly string[]): ComparisonUnitGroup {
  const words: ComparisonUnitWord[] = wordHashes.map((hash) => makeWord(hash));
  const paragraph: ComparisonUnitGroup = {
    kind: 'paragraph',
    sha1Hash: `p:${wordHashes.join('|')}`,
    contents: words,
  };
  const cell: ComparisonUnitGroup = {
    kind: 'cell',
    sha1Hash: `c:${wordHashes.join('|')}`,
    contents: [paragraph],
  };
  return {
    kind: 'row',
    sha1Hash: `r:${wordHashes.join('|')}`,
    contents: [cell],
  };
}

function unknownRows(
  left: readonly (readonly string[])[],
  right: readonly (readonly string[])[],
): CorrelatedSequence {
  return {
    correlationStatus: 'unknown',
    comparisonUnits1: left.map(makeRow),
    comparisonUnits2: right.map(makeRow),
  };
}

function countUnits(units: readonly ComparisonUnit[]): number {
  return units.length;
}

describe('findCommonAtBeginningAndEnd', () => {
  it('extracts equal prefix and suffix around unknown middle', () => {
    const result = findCommonAtBeginningAndEnd(
      unknown(['A', 'B', 'C', 'D'], ['A', 'X', 'C', 'D']),
    );

    expect(statuses(result)).toEqual(['equal', 'unknown', 'equal']);
    expect(result[0]!.comparisonUnits1).toHaveLength(1);
    expect(result[1]!.comparisonUnits1).toHaveLength(1);
    expect(result[1]!.comparisonUnits2).toHaveLength(1);
    expect(result[2]!.comparisonUnits1).toHaveLength(2);
  });

  it('returns inserted when left side is empty after trimming', () => {
    const result = findCommonAtBeginningAndEnd(unknown(['A'], ['A', 'B']));
    expect(statuses(result)).toEqual(['equal', 'inserted']);
  });

  it('returns deleted when right side is empty after trimming', () => {
    const result = findCommonAtBeginningAndEnd(unknown(['A', 'B'], ['A']));
    expect(statuses(result)).toEqual(['equal', 'deleted']);
  });
});

describe('doLcsAlgorithm', () => {
  it('returns one equal segment for identical sequences', () => {
    const result = doLcsAlgorithm(unknown(['A', 'B', 'C'], ['A', 'B', 'C']));

    expect(statuses(result)).toEqual(['equal']);
    expect(result[0]!.comparisonUnits1).toHaveLength(3);
    expect(result[0]!.comparisonUnits2).toHaveLength(3);
  });

  it('returns inserted when baseline has no units', () => {
    const result = doLcsAlgorithm(unknown([], ['A', 'B']));

    expect(statuses(result)).toEqual(['inserted']);
    expect(result[0]!.comparisonUnits1).toHaveLength(0);
    expect(result[0]!.comparisonUnits2).toHaveLength(2);
  });

  it('returns deleted when candidate has no units', () => {
    const result = doLcsAlgorithm(unknown(['A', 'B'], []));

    expect(statuses(result)).toEqual(['deleted']);
    expect(result[0]!.comparisonUnits1).toHaveLength(2);
    expect(result[0]!.comparisonUnits2).toHaveLength(0);
  });
});

describe('doLcsAlgorithmForTable', () => {
  it('matches rows with partial word overlap as equal anchors', () => {
    const result = doLcsAlgorithmForTable(
      unknownRows(
        [
          ['A', 'B'],
          ['C', 'D'],
        ],
        [
          ['A', 'X'],
          ['C', 'D'],
        ],
      ),
    );

    expect(statuses(result)).toEqual(['equal']);
    expect(countUnits(result[0]!.comparisonUnits1)).toBe(2);
    expect(countUnits(result[0]!.comparisonUnits2)).toBe(2);
  });

  it('falls back to inserted/deleted when rows are dissimilar', () => {
    const result = doLcsAlgorithmForTable(
      unknownRows([
        ['LEFT'],
      ], [
        ['RIGHT'],
      ]),
    );

    expect(statuses(result)).toEqual(['deleted', 'inserted']);
  });
});

describe('processCorrelatedHashes', () => {
  it('anchors unique matching hashes as equal and leaves surrounding unknown spans', () => {
    const result = processCorrelatedHashes(
      unknown(['X', 'A', 'B', 'Y'], ['Q', 'A', 'B', 'R']),
    );

    expect(statuses(result)).toEqual(['unknown', 'equal', 'unknown']);
    expect(result[0]!.comparisonUnits1).toHaveLength(1);
    expect(result[0]!.comparisonUnits2).toHaveLength(1);
    expect(result[1]!.comparisonUnits1.map((u) => u.sha1Hash)).toEqual(['A', 'B']);
  });

  it('ignores duplicated hashes but still anchors other unique hashes', () => {
    const result = processCorrelatedHashes(
      unknown(['A', 'A', 'B'], ['A', 'A', 'B']),
    );

    expect(statuses(result)).toEqual(['unknown', 'equal']);
    expect(result[1]!.comparisonUnits1[0]!.sha1Hash).toBe('B');
  });

  it('uses only monotonic anchors when right-side positions cross', () => {
    const result = processCorrelatedHashes(
      unknown(['A', 'B'], ['B', 'A']),
    );

    expect(statuses(result)).toEqual(['inserted', 'equal', 'deleted']);
    expect(result[1]!.comparisonUnits1[0]!.sha1Hash).toBe('A');
  });
});

describe('correlateComparisonUnits', () => {
  it('splits a single replacement into deleted + inserted with equal context', () => {
    const result = correlateComparisonUnits(
      ['A', 'B', 'C', 'D'].map(makeWord),
      ['A', 'X', 'C', 'D'].map(makeWord),
    );

    expect(statuses(result)).toEqual(['equal', 'deleted', 'inserted', 'equal']);
    expect(result[0]!.comparisonUnits1).toHaveLength(1);
    expect(result[1]!.comparisonUnits1).toHaveLength(1);
    expect(result[2]!.comparisonUnits2).toHaveLength(1);
    expect(result[3]!.comparisonUnits1).toHaveLength(2);
  });

  it('coalesces full equality into one segment', () => {
    const result = correlateComparisonUnits(
      ['A', 'B'].map(makeWord),
      ['A', 'B'].map(makeWord),
    );

    expect(statuses(result)).toEqual(['equal']);
    expect(result[0]!.comparisonUnits1).toHaveLength(2);
  });

  it('produces deterministic status sequence for transposition case', () => {
    const left = ['A', 'B', 'C'].map(makeWord);
    const right = ['B', 'A', 'C'].map(makeWord);

    const run1 = correlateComparisonUnits(left, right);
    const run2 = correlateComparisonUnits(left, right);

    expect(statuses(run1)).toEqual(statuses(run2));
    expect(statuses(run1)).toEqual(['inserted', 'equal', 'deleted', 'equal']);
  });

  it('resolves long middle spans via hash anchors before LCS', () => {
    const result = correlateComparisonUnits(
      ['start', 'A', 'left-only', 'B', 'end'].map(makeWord),
      ['start', 'A', 'right-only', 'B', 'end'].map(makeWord),
    );

    expect(statuses(result)).toEqual(['equal', 'deleted', 'inserted', 'equal']);
  });

  it('uses table-row LCS to keep similar rows aligned as equal segments', () => {
    const result = correlateComparisonUnits(
      [
        ['A', 'B'],
        ['C', 'D'],
      ].map(makeRow),
      [
        ['A', 'X'],
        ['C', 'D'],
      ].map(makeRow),
    );

    expect(statuses(result)).toEqual(['equal']);
    expect(result[0]!.comparisonUnits1).toHaveLength(2);
    expect(result[0]!.comparisonUnits2).toHaveLength(2);
  });
});
