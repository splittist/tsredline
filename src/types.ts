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
}

export interface CompareResult {
  readonly equal: boolean;
  readonly summary: string;
  readonly changes: readonly ComparisonChange[];
  readonly notices: readonly ComparisonNotice[];
  readonly metadata: ComparisonMetadata;
}
