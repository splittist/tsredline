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

function splitUnknownRange(
  left: readonly ComparisonUnit[],
  right: readonly ComparisonUnit[],
): readonly CorrelatedSequence[] {
  if (left.length > 0 && right.length > 0) {
    return [
      {
        correlationStatus: 'unknown',
        comparisonUnits1: left,
        comparisonUnits2: right,
      },
    ];
  }

  if (left.length > 0) {
    return [
      {
        correlationStatus: 'deleted',
        comparisonUnits1: left,
        comparisonUnits2: [],
      },
    ];
  }

  if (right.length > 0) {
    return [
      {
        correlationStatus: 'inserted',
        comparisonUnits1: [],
        comparisonUnits2: right,
      },
    ];
  }

  return [];
}

export function processCorrelatedHashes(
  unknown: CorrelatedSequence,
): readonly CorrelatedSequence[] {
  if (unknown.correlationStatus !== 'unknown') {
    return [unknown];
  }

  const left = unknown.comparisonUnits1;
  const right = unknown.comparisonUnits2;

  const leftPos = new Map<string, number[]>();
  const rightPos = new Map<string, number[]>();

  for (let i = 0; i < left.length; i += 1) {
    const hash = left[i]!.sha1Hash;
    const arr = leftPos.get(hash) ?? [];
    arr.push(i);
    leftPos.set(hash, arr);
  }

  for (let i = 0; i < right.length; i += 1) {
    const hash = right[i]!.sha1Hash;
    const arr = rightPos.get(hash) ?? [];
    arr.push(i);
    rightPos.set(hash, arr);
  }

  const candidates: Array<{ leftIndex: number; rightIndex: number }> = [];
  for (const [hash, leftIndexes] of leftPos) {
    if (leftIndexes.length !== 1) continue;
    const rightIndexes = rightPos.get(hash);
    if (rightIndexes === undefined || rightIndexes.length !== 1) continue;
    candidates.push({
      leftIndex: leftIndexes[0]!,
      rightIndex: rightIndexes[0]!,
    });
  }

  if (candidates.length === 0) {
    return [unknown];
  }

  candidates.sort((a, b) => a.leftIndex - b.leftIndex);

  // Keep anchors in monotonic right-side order to avoid crossing matches.
  const anchors: Array<{ leftIndex: number; rightIndex: number }> = [];
  let lastRight = -1;
  for (const c of candidates) {
    if (c.rightIndex > lastRight) {
      anchors.push(c);
      lastRight = c.rightIndex;
    }
  }

  if (anchors.length === 0) {
    return [unknown];
  }

  const result: CorrelatedSequence[] = [];
  let cursorLeft = 0;
  let cursorRight = 0;

  for (const anchor of anchors) {
    const beforeLeft = left.slice(cursorLeft, anchor.leftIndex);
    const beforeRight = right.slice(cursorRight, anchor.rightIndex);
    for (const seg of splitUnknownRange(beforeLeft, beforeRight)) {
      pushOrMerge(result, seg);
    }

    pushOrMerge(result, {
      correlationStatus: 'equal',
      comparisonUnits1: [left[anchor.leftIndex]!],
      comparisonUnits2: [right[anchor.rightIndex]!],
    });

    cursorLeft = anchor.leftIndex + 1;
    cursorRight = anchor.rightIndex + 1;
  }

  for (const seg of splitUnknownRange(
    left.slice(cursorLeft),
    right.slice(cursorRight),
  )) {
    pushOrMerge(result, seg);
  }

  return result.length > 0 ? result : [unknown];
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
      const fastPathSegments = processCorrelatedHashes(segment);
      for (const fastSegment of fastPathSegments) {
        if (fastSegment.correlationStatus !== 'unknown') {
          pushOrMerge(resolved, fastSegment);
          continue;
        }

        const retrimmed = findCommonAtBeginningAndEnd(fastSegment);
        for (const retrimmedSegment of retrimmed) {
          if (retrimmedSegment.correlationStatus === 'unknown') {
            for (const lcsSegment of doLcsAlgorithm(retrimmedSegment)) {
              pushOrMerge(resolved, lcsSegment);
            }
          } else {
            pushOrMerge(resolved, retrimmedSegment);
          }
        }
      }
    } else {
      pushOrMerge(resolved, segment);
    }
  }

  return resolved;
}