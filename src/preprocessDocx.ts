import JSZip from 'jszip';

export type RevisionMode = 'accept' | 'reject' | 'preserve';

export interface PreprocessDocxOptions {
  readonly ignoreWhitespace?: boolean;
  readonly revisionMode?: RevisionMode;
}

export interface PreprocessedDocx {
  readonly normalizedText: string;
  readonly paragraphCount: number;
  readonly wordCount: number;
}

function getLocalName(node: Node): string {
  return node.nodeName.includes(':')
    ? node.nodeName.slice(node.nodeName.indexOf(':') + 1)
    : node.nodeName;
}

function cloneAttribute(target: Element, attribute: Attr): void {
  if (attribute.namespaceURI) {
    target.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
    return;
  }
  target.setAttribute(attribute.name, attribute.value);
}

function cloneAsWordText(document: XMLDocument, element: Element): Element {
  const replacement = document.createElementNS(element.namespaceURI, 'w:t');
  for (const attribute of Array.from(element.attributes)) {
    cloneAttribute(replacement, attribute);
  }
  replacement.textContent = element.textContent;
  return replacement;
}

function transformRevisionNode(
  node: Node,
  document: XMLDocument,
  revisionMode: RevisionMode,
): Node[] {
  if (node.nodeType === Node.TEXT_NODE) {
    return [document.createTextNode(node.textContent ?? '')];
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return [];
  }

  const element = node as Element;
  const localName = getLocalName(element);

  if (localName === 'ins' || localName === 'moveTo') {
    if (revisionMode === 'reject') {
      return [];
    }

    return Array.from(element.childNodes).flatMap((child) =>
      transformRevisionNode(child, document, revisionMode),
    );
  }

  if (localName === 'del' || localName === 'moveFrom') {
    if (revisionMode === 'accept') {
      return [];
    }

    return Array.from(element.childNodes).flatMap((child) =>
      transformRevisionNode(child, document, revisionMode),
    );
  }

  if (localName === 'delText' && revisionMode === 'reject') {
    return [cloneAsWordText(document, element)];
  }

  const clonedElement = document.createElementNS(element.namespaceURI, element.nodeName);
  for (const attribute of Array.from(element.attributes)) {
    cloneAttribute(clonedElement, attribute);
  }
  for (const child of Array.from(element.childNodes)) {
    const transformedChildren = transformRevisionNode(child, document, revisionMode);
    for (const transformedChild of transformedChildren) {
      clonedElement.appendChild(transformedChild);
    }
  }
  return [clonedElement];
}

export function preprocessRevisionMarkup(
  documentXml: string,
  revisionMode: RevisionMode,
): string {
  if (revisionMode === 'preserve') {
    return documentXml;
  }

  const parser = new DOMParser();
  const parsed = parser.parseFromString(documentXml, 'application/xml');
  const root = parsed.documentElement;
  if (!root || root.getElementsByTagName('parsererror').length > 0) {
    throw new Error('Failed to parse word/document.xml for revision preprocessing.');
  }

  const output = document.implementation.createDocument(root.namespaceURI, root.nodeName);
  const outputRoot = output.documentElement;
  for (const attribute of Array.from(root.attributes)) {
    cloneAttribute(outputRoot, attribute);
  }
  outputRoot.textContent = '';

  for (const child of Array.from(root.childNodes)) {
    const transformedChildren = transformRevisionNode(child, output, revisionMode);
    for (const transformedChild of transformedChildren) {
      outputRoot.appendChild(transformedChild);
    }
  }

  return new XMLSerializer().serializeToString(output);
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
  options: PreprocessDocxOptions = {},
): Promise<PreprocessedDocx> {
  const zip = await JSZip.loadAsync(new Uint8Array(docx));
  const documentPart = zip.file('word/document.xml');
  if (!documentPart) {
    throw new Error('DOCX part not found: word/document.xml');
  }

  const documentXml = await documentPart.async('text');
  const revisionMode = options.revisionMode ?? 'accept';
  const preprocessedXml = preprocessRevisionMarkup(documentXml, revisionMode);
  const paragraphCount = preprocessedXml.match(/<w:p\b[^>]*>/g)?.length ?? 0;

  const rawText = extractDocumentText(preprocessedXml);
  const normalizedText = options.ignoreWhitespace
    ? normalizeWhitespace(rawText)
    : rawText;

  return {
    normalizedText,
    paragraphCount,
    wordCount: countWords(normalizedText),
  };
}