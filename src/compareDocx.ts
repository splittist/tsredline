import type { CompareOptions, CompareResult, ComparisonChange } from './types';
import { assignUnids } from './assignUnids';
import { correlateComparisonUnits } from './correlateComparisonUnits';
import { createAtomList } from './createAtomList';
import { getComparisonUnitList } from './getComparisonUnitList';
import { preprocessDocx } from './preprocessDocx';
import type { ComparisonUnitWord } from './getComparisonUnitList';
import { arrayBuffersEqual } from './utils/arrayBuffer';

const SKELETON_NOTICE =
  'DOCX structural comparison is not implemented yet. This scaffold only validates inputs and reports coarse binary-level differences.';
const PREPROCESS_NOTICE =
  'Phase-1 preprocessing is active: comparison currently uses normalized word/document.xml text while deeper structural phases are still pending.';
const COMPARISON_UNIT_NOTICE =
  'Comparison-unit correlation is active: assignUnids/createAtomList/getComparisonUnitList outputs are correlated to provide structural unit counts.';

function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === '[object ArrayBuffer]';
}

function assertArrayBuffer(value: ArrayBuffer, label: string): void {
  if (!isArrayBuffer(value)) {
    throw new TypeError(`${label} must be an ArrayBuffer.`);
  }
}

function wordsFromText(text: string): readonly string[] {
  return text.split(/\s+/).filter((w) => w.length > 0);
}

function toWordUnits(words: readonly string[]): readonly ComparisonUnitWord[] {
  return words.map((word) => ({
    kind: 'word',
    sha1Hash: word,
    atoms: [],
  }));
}

export async function compareDocx(
  baseline: ArrayBuffer,
  candidate: ArrayBuffer,
  options: CompareOptions = {},
): Promise<CompareResult> {
  assertArrayBuffer(baseline, 'baseline');
  assertArrayBuffer(candidate, 'candidate');

  const identicalBinary = arrayBuffersEqual(baseline, candidate);

  try {
    const preprocessOptions =
      options.ignoreWhitespace === undefined
        ? {}
        : { ignoreWhitespace: options.ignoreWhitespace };

    const baselinePreprocessed = await preprocessDocx(
      baseline,
      preprocessOptions,
    );
    const candidatePreprocessed = await preprocessDocx(
      candidate,
      preprocessOptions,
    );

    const equal =
      baselinePreprocessed.normalizedText === candidatePreprocessed.normalizedText;

    const baselineWords = wordsFromText(baselinePreprocessed.normalizedText);
    const candidateWords = wordsFromText(candidatePreprocessed.normalizedText);
    const correlatedTextUnits = correlateComparisonUnits(
      toWordUnits(baselineWords),
      toWordUnits(candidateWords),
    );

    let equalUnits = 0;
    let deletedUnits = 0;
    let insertedUnits = 0;
    for (const segment of correlatedTextUnits) {
      if (segment.correlationStatus === 'equal') {
        equalUnits += segment.comparisonUnits1.length;
      } else if (segment.correlationStatus === 'deleted') {
        deletedUnits += segment.comparisonUnits1.length;
      } else if (segment.correlationStatus === 'inserted') {
        insertedUnits += segment.comparisonUnits2.length;
      }
    }

    const baselineStamped = await assignUnids(baseline);
    const candidateStamped = await assignUnids(candidate);
    const baselineAtoms = await createAtomList(baselineStamped.docx);
    const candidateAtoms = await createAtomList(candidateStamped.docx);
    const baselineComparisonUnits = await getComparisonUnitList(baselineAtoms.atoms);
    const candidateComparisonUnits = await getComparisonUnitList(candidateAtoms.atoms);
    const correlatedComparisonUnits = correlateComparisonUnits(
      baselineComparisonUnits,
      candidateComparisonUnits,
    );

    let equalComparisonUnits = 0;
    let deletedComparisonUnits = 0;
    let insertedComparisonUnits = 0;
    for (const segment of correlatedComparisonUnits) {
      if (segment.correlationStatus === 'equal') {
        equalComparisonUnits += segment.comparisonUnits1.length;
      } else if (segment.correlationStatus === 'deleted') {
        deletedComparisonUnits += segment.comparisonUnits1.length;
      } else if (segment.correlationStatus === 'inserted') {
        insertedComparisonUnits += segment.comparisonUnits2.length;
      }
    }

    return {
      equal,
      summary: equal
        ? 'Preprocessed document text matches. Deeper XML-structure diff phases are still pending.'
        : 'Preprocessed document text differs. Deeper XML-structure diff phases are still pending.',
      changes: equal
        ? []
        : (() => {
            const nextChanges: ComparisonChange[] = [
              {
                kind: 'replace',
                path: 'word/document.xml:text',
                before: `${baselinePreprocessed.wordCount} words in ${baselinePreprocessed.paragraphCount} paragraphs`,
                after: `${candidatePreprocessed.wordCount} words in ${candidatePreprocessed.paragraphCount} paragraphs`,
              },
            ];
            if (deletedUnits > 0) {
              nextChanges.push({
                kind: 'delete',
                path: 'word/document.xml:correlation',
                before: `${deletedUnits} preprocessed unit(s)`,
              });
            }
            if (insertedUnits > 0) {
              nextChanges.push({
                kind: 'insert',
                path: 'word/document.xml:correlation',
                after: `${insertedUnits} preprocessed unit(s)`,
              });
            }
            if (deletedComparisonUnits > 0) {
              nextChanges.push({
                kind: 'delete',
                path: 'word/document.xml:comparison-units',
                before: `${deletedComparisonUnits} comparison unit(s)`,
              });
            }
            if (insertedComparisonUnits > 0) {
              nextChanges.push({
                kind: 'insert',
                path: 'word/document.xml:comparison-units',
                after: `${insertedComparisonUnits} comparison unit(s)`,
              });
            }
            return nextChanges;
          })(),
      notices: [
        {
          code: 'SKELETON_ENGINE',
          message: options.includeDiagnostics
            ? `${SKELETON_NOTICE} ${PREPROCESS_NOTICE} ${COMPARISON_UNIT_NOTICE}`
            : SKELETON_NOTICE,
        },
      ],
      metadata: {
        baselineSize: baseline.byteLength,
        candidateSize: candidate.byteLength,
        identicalBinary,
        comparisonMode: 'preprocessed-text',
        baselineParagraphs: baselinePreprocessed.paragraphCount,
        candidateParagraphs: candidatePreprocessed.paragraphCount,
        baselineUnits: baselineWords.length,
        candidateUnits: candidateWords.length,
        equalUnits,
        deletedUnits,
        insertedUnits,
        baselineComparisonUnits: baselineComparisonUnits.length,
        candidateComparisonUnits: candidateComparisonUnits.length,
        equalComparisonUnits,
        deletedComparisonUnits,
        insertedComparisonUnits,
      },
    };
  } catch {
    // Keep binary fallback behavior for non-DOCX inputs used in unit tests or diagnostics.
  }

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
      comparisonMode: 'binary-fallback',
    },
  };
}
