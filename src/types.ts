export interface CompareOptions {
  readonly ignoreWhitespace?: boolean;
  readonly trackMoves?: boolean;
  readonly includeDiagnostics?: boolean;
}

export interface ComparisonChange {
  readonly kind: 'insert' | 'delete' | 'replace';
  readonly path: string;
  readonly before?: string;
  readonly after?: string;
}

export interface ComparisonNotice {
  readonly code: 'SKELETON_ENGINE';
  readonly message: string;
}

export interface ComparisonMetadata {
  readonly baselineSize: number;
  readonly candidateSize: number;
  readonly identicalBinary: boolean;
  readonly comparisonMode?: 'preprocessed-text' | 'binary-fallback';
  readonly baselineParagraphs?: number;
  readonly candidateParagraphs?: number;
  readonly baselineUnits?: number;
  readonly candidateUnits?: number;
  readonly equalUnits?: number;
  readonly deletedUnits?: number;
  readonly insertedUnits?: number;
  readonly baselineComparisonUnits?: number;
  readonly candidateComparisonUnits?: number;
  readonly equalComparisonUnits?: number;
  readonly deletedComparisonUnits?: number;
  readonly insertedComparisonUnits?: number;
}

export interface CompareResult {
  readonly equal: boolean;
  readonly summary: string;
  readonly changes: readonly ComparisonChange[];
  readonly notices: readonly ComparisonNotice[];
  readonly metadata: ComparisonMetadata;
}
