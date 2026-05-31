import type { CompareOptions, CompareResult } from './types';
import { arrayBuffersEqual } from './utils/arrayBuffer';

const SKELETON_NOTICE =
  'DOCX structural comparison is not implemented yet. This scaffold only validates inputs and reports coarse binary-level differences.';

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === '[object ArrayBuffer]';
}

function assertArrayBuffer(value: ArrayBuffer, label: string): void {
  if (!isArrayBuffer(value)) {
    throw new TypeError(`${label} must be an ArrayBuffer.`);
  }
}

export async function compareDocx(
  baseline: ArrayBuffer,
  candidate: ArrayBuffer,
  options: CompareOptions = {},
): Promise<CompareResult> {
  assertArrayBuffer(baseline, 'baseline');
  assertArrayBuffer(candidate, 'candidate');

  const identicalBinary = arrayBuffersEqual(baseline, candidate);

  return {
    equal: identicalBinary,
    summary: identicalBinary
      ? 'Inputs are byte-identical. The DOCX-aware comparison pipeline is ready to be implemented on top of this scaffold.'
      : 'Inputs differ at the binary level. The DOCX-aware comparison pipeline is ready to be implemented on top of this scaffold.',
    changes: identicalBinary
      ? []
      : [
          {
            kind: 'replace',
            path: 'docx-package',
            before: `${baseline.byteLength} bytes`,
            after: `${candidate.byteLength} bytes`,
          },
        ],
    notices: [
      {
        code: 'SKELETON_ENGINE',
        message: options.includeDiagnostics
          ? `${SKELETON_NOTICE} Diagnostic output hooks are reserved for the future document pipeline.`
          : SKELETON_NOTICE,
      },
    ],
    metadata: {
      baselineSize: baseline.byteLength,
      candidateSize: candidate.byteLength,
      identicalBinary,
    },
  };
}
