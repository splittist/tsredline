import JSZip from 'jszip';

export interface PreprocessedDocx {
  readonly normalizedText: string;
  readonly paragraphCount: number;
  readonly wordCount: number;
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

function extractDocumentText(documentXml: string): string {
  const textChunks: string[] = [];
  const tokenPattern = /<w:p\b[^>]*>|<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;

  let match = tokenPattern.exec(documentXml);
  while (match) {
    if (match[0].startsWith('<w:p')) {
      textChunks.push('\n');
    } else if (typeof match[1] === 'string' && match[1].length > 0) {
      textChunks.push(decodeXmlEntities(match[1]));
    }
    match = tokenPattern.exec(documentXml);
  }

  return textChunks.join('');
}

function countWords(text: string): number {
  if (text.length === 0) {
    return 0;
  }
  return text.split(/\s+/).filter((token) => token.length > 0).length;
}

export async function preprocessDocx(
  docx: ArrayBuffer,
  options: { readonly ignoreWhitespace?: boolean } = {},
): Promise<PreprocessedDocx> {
  const zip = await JSZip.loadAsync(new Uint8Array(docx));
  const documentPart = zip.file('word/document.xml');
  if (!documentPart) {
    throw new Error('DOCX part not found: word/document.xml');
  }

  const documentXml = await documentPart.async('text');
  const paragraphCount = documentXml.match(/<w:p\b[^>]*>/g)?.length ?? 0;

  const rawText = extractDocumentText(documentXml);
  const normalizedText = options.ignoreWhitespace
    ? normalizeWhitespace(rawText)
    : rawText;

  return {
    normalizedText,
    paragraphCount,
    wordCount: countWords(normalizedText),
  };
}