import JSZip from 'jszip';

import { collectContentParts } from './docxParts';
import { preprocessRevisionMarkup, type RevisionMode } from './preprocessDocx';

export interface HashBlockLevelOptions {
  readonly ignoreWhitespace?: boolean;
  readonly includeTableRows?: boolean;
  readonly minimumWordCount?: number;
  readonly revisionMode?: RevisionMode;
}

export interface BlockHash {
  readonly kind: 'paragraph' | 'table-row';
  readonly partName: string;
  readonly index: number;
  readonly wordCount: number;
  readonly hash: string;
  readonly text: string;
}

export interface HashBlockLevelResult {
  readonly blocks: readonly BlockHash[];
  readonly blockCount: number;
}

function decodeXmlEntities(value: string): string {
  return value
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&');
}

function normalizeWhitespace(value: string): string {
  return value.replace(/\s+/g, ' ').trim();
}

function countWords(text: string): number {
  if (text.length === 0) {
    return 0;
  }
  return text.split(/\s+/).filter((token) => token.length > 0).length;
}

function extractBlockText(blockXml: string, ignoreWhitespace: boolean): string {
  const pieces: string[] = [];
  const textPattern = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
  let textMatch = textPattern.exec(blockXml);
  while (textMatch) {
    const raw = textMatch[1] ?? '';
    if (raw.length > 0) {
      pieces.push(decodeXmlEntities(raw));
    }
    textMatch = textPattern.exec(blockXml);
  }

  const joined = pieces.join(' ');
  return ignoreWhitespace ? normalizeWhitespace(joined) : joined;
}

function bytesToHex(buffer: ArrayBuffer): string {
  const byteArray = new Uint8Array(buffer);
  let output = '';
  for (const value of byteArray) {
    output += value.toString(16).padStart(2, '0');
  }
  return output;
}

async function sha1Hex(value: string): Promise<string> {
  if (!globalThis.crypto?.subtle) {
    throw new Error('Web Crypto API is not available for block hashing.');
  }
  const bytes = new TextEncoder().encode(value);
  const digest = await globalThis.crypto.subtle.digest('SHA-1', bytes);
  return bytesToHex(digest);
}

function collectBlocks(
  partXml: string,
  options: Required<Pick<HashBlockLevelOptions, 'includeTableRows' | 'ignoreWhitespace'>>,
): Array<{ kind: 'paragraph' | 'table-row'; text: string }> {
  const results: Array<{ kind: 'paragraph' | 'table-row'; text: string }> = [];
  const paragraphPattern = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  let paragraphMatch = paragraphPattern.exec(partXml);
  while (paragraphMatch) {
    const text = extractBlockText(paragraphMatch[0], options.ignoreWhitespace);
    results.push({ kind: 'paragraph', text });
    paragraphMatch = paragraphPattern.exec(partXml);
  }

  if (options.includeTableRows) {
    const rowPattern = /<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g;
    let rowMatch = rowPattern.exec(partXml);
    while (rowMatch) {
      const text = extractBlockText(rowMatch[0], options.ignoreWhitespace);
      results.push({ kind: 'table-row', text });
      rowMatch = rowPattern.exec(partXml);
    }
  }

  return results;
}

export async function hashBlockLevelContent(
  docx: ArrayBuffer,
  options: HashBlockLevelOptions = {},
): Promise<HashBlockLevelResult> {
  const includeTableRows = options.includeTableRows ?? false;
  const ignoreWhitespace = options.ignoreWhitespace ?? true;
  const minimumWordCount = options.minimumWordCount ?? 0;
  const revisionMode = options.revisionMode ?? 'accept';

  const zip = await JSZip.loadAsync(new Uint8Array(docx));
  if (!zip.file('word/document.xml')) {
    throw new Error('DOCX part not found: word/document.xml');
  }

  const hashedBlocks: BlockHash[] = [];
  for (const partName of collectContentParts(zip)) {
    const partFile = zip.file(partName);
    if (!partFile) {
      continue;
    }

    const partXml = await partFile.async('text');
    const preprocessedXml = preprocessRevisionMarkup(partXml, revisionMode);
    const blocks = collectBlocks(preprocessedXml, {
      includeTableRows,
      ignoreWhitespace,
    });

    for (let index = 0; index < blocks.length; index += 1) {
      const block = blocks[index];
      if (!block) {
        continue;
      }
      const wordCount = countWords(block.text);
      if (wordCount < minimumWordCount) {
        continue;
      }

      hashedBlocks.push({
        kind: block.kind,
        partName,
        index,
        wordCount,
        text: block.text,
        hash: await sha1Hex(block.text),
      });
    }
  }

  return {
    blocks: hashedBlocks,
    blockCount: hashedBlocks.length,
  };
}