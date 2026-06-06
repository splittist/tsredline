import type { ComparisonUnit } from './getComparisonUnitList';

export type CorrelationStatus = 'unknown' | 'equal' | 'deleted' | 'inserted';

export interface CorrelatedSequence {
  readonly correlationStatus: CorrelationStatus;
  readonly comparisonUnits1: readonly ComparisonUnit[];
  readonly comparisonUnits2: readonly ComparisonUnit[];
}

function hashesEqual(left: ComparisonUnit, right: ComparisonUnit): boolean {
  return left.sha1Hash === right.sha1Hash;
}

function pushOrMerge(
  result: CorrelatedSequence[],
  next: CorrelatedSequence,
): void {
  const last = result[result.length - 1];
  if (last === undefined || last.correlationStatus !== next.correlationStatus) {
    result.push(next);
    return;
  }

  result[result.length - 1] = {
    correlationStatus: last.correlationStatus,
    comparisonUnits1: [...last.comparisonUnits1, ...next.comparisonUnits1],
    comparisonUnits2: [...last.comparisonUnits2, ...next.comparisonUnits2],
  };
}

export function findCommonAtBeginningAndEnd(
  unknown: CorrelatedSequence,
): readonly CorrelatedSequence[] {
  const left = unknown.comparisonUnits1;
  const right = unknown.comparisonUnits2;

  if (unknown.correlationStatus !== 'unknown') {
    return [unknown];
  }

  let prefix = 0;
  while (
    prefix < left.length &&
    prefix < right.length &&
    hashesEqual(left[prefix]!, right[prefix]!)
  ) {
    prefix += 1;
  }

  let suffix = 0;
  while (
    suffix < left.length - prefix &&
    suffix < right.length - prefix &&
    hashesEqual(left[left.length - 1 - suffix]!, right[right.length - 1 - suffix]!)
  ) {
    suffix += 1;
  }

  const middleLeft = left.slice(prefix, left.length - suffix);
  const middleRight = right.slice(prefix, right.length - suffix);
  const result: CorrelatedSequence[] = [];

  if (prefix > 0) {
    result.push({
      correlationStatus: 'equal',
      comparisonUnits1: left.slice(0, prefix),
      comparisonUnits2: right.slice(0, prefix),
    });
  }

  if (middleLeft.length > 0 || middleRight.length > 0) {
    if (middleLeft.length === 0) {
      result.push({
        correlationStatus: 'inserted',
        comparisonUnits1: [],
        comparisonUnits2: middleRight,
      });
    } else if (middleRight.length === 0) {
      result.push({
        correlationStatus: 'deleted',
        comparisonUnits1: middleLeft,
        comparisonUnits2: [],
      });
    } else {
      result.push({
        correlationStatus: 'unknown',
        comparisonUnits1: middleLeft,
        comparisonUnits2: middleRight,
      });
    }
  }

  if (suffix > 0) {
    result.push({
      correlationStatus: 'equal',
      comparisonUnits1: left.slice(left.length - suffix),
      comparisonUnits2: right.slice(right.length - suffix),
    });
  }

  return result;
}

export function doLcsAlgorithm(
  unknown: CorrelatedSequence,
): readonly CorrelatedSequence[] {
  const left = unknown.comparisonUnits1;
  const right = unknown.comparisonUnits2;

  if (unknown.correlationStatus !== 'unknown') {
    return [unknown];
  }

  if (left.length === 0 && right.length === 0) {
    return [];
  }

  if (left.length === 0) {
    return [
      {
        correlationStatus: 'inserted',
        comparisonUnits1: [],
        comparisonUnits2: right,
      },
    ];
  }

  if (right.length === 0) {
    return [
      {
        correlationStatus: 'deleted',
        comparisonUnits1: left,
        comparisonUnits2: [],
      },
    ];
  }

  const rows = left.length + 1;
  const cols = right.length + 1;
  const dp: number[][] = Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => 0),
  );

  for (let i = 1; i < rows; i += 1) {
    for (let j = 1; j < cols; j += 1) {
      if (hashesEqual(left[i - 1]!, right[j - 1]!)) {
        dp[i]![j] = (dp[i - 1]![j - 1] ?? 0) + 1;
      } else {
        dp[i]![j] = Math.max(dp[i - 1]![j] ?? 0, dp[i]![j - 1] ?? 0);
      }
    }
  }

  const reversed: CorrelatedSequence[] = [];
  let i = left.length;
  let j = right.length;

  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && hashesEqual(left[i - 1]!, right[j - 1]!)) {
      pushOrMerge(reversed, {
        correlationStatus: 'equal',
        comparisonUnits1: [left[i - 1]!],
        comparisonUnits2: [right[j - 1]!],
      });
      i -= 1;
      j -= 1;
      continue;
    }

    const up = i > 0 ? (dp[i - 1]![j] ?? 0) : -1;
    const leftward = j > 0 ? (dp[i]![j - 1] ?? 0) : -1;
    if (i > 0 && (j === 0 || up > leftward)) {
      pushOrMerge(reversed, {
        correlationStatus: 'deleted',
        comparisonUnits1: [left[i - 1]!],
        comparisonUnits2: [],
      });
      i -= 1;
    } else if (j > 0) {
      pushOrMerge(reversed, {
        correlationStatus: 'inserted',
        comparisonUnits1: [],
        comparisonUnits2: [right[j - 1]!],
      });
      j -= 1;
    }
  }

  return reversed.reverse().map((segment) => ({
    correlationStatus: segment.correlationStatus,
    comparisonUnits1: [...segment.comparisonUnits1].reverse(),
    comparisonUnits2: [...segment.comparisonUnits2].reverse(),
  }));
}

export function correlateComparisonUnits(
  comparisonUnits1: readonly ComparisonUnit[],
  comparisonUnits2: readonly ComparisonUnit[],
): readonly CorrelatedSequence[] {
  const initial: CorrelatedSequence = {
    correlationStatus: 'unknown',
    comparisonUnits1,
    comparisonUnits2,
  };

  const trimmed = findCommonAtBeginningAndEnd(initial);
  const resolved: CorrelatedSequence[] = [];

  for (const segment of trimmed) {
    if (segment.correlationStatus === 'unknown') {
      for (const lcsSegment of doLcsAlgorithm(segment)) {
        pushOrMerge(resolved, lcsSegment);
      }
    } else {
      pushOrMerge(resolved, segment);
    }
  }

  return resolved;
}