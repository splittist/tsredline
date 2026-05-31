import JSZip from 'jszip';

function toZipInput(docx: ArrayBuffer): Uint8Array {
  return new Uint8Array(docx);
}

export async function loadDocxZip(docx: ArrayBuffer): Promise<JSZip> {
  return JSZip.loadAsync(toZipInput(docx));
}

export async function listDocxParts(docx: ArrayBuffer): Promise<string[]> {
  const zip = await loadDocxZip(docx);
  return Object.keys(zip.files).sort();
}

export async function readDocxPartText(
  docx: ArrayBuffer,
  partName: string,
): Promise<string> {
  const zip = await loadDocxZip(docx);
  const part = zip.file(partName);
  if (!part) {
    throw new Error(`DOCX part not found: ${partName}`);
  }
  return part.async('text');
}

export async function hasDocxPart(
  docx: ArrayBuffer,
  partName: string,
): Promise<boolean> {
  const zip = await loadDocxZip(docx);
  return zip.file(partName) !== null;
}

export async function countXmlTagInPart(
  docx: ArrayBuffer,
  partName: string,
  tagName: string,
): Promise<number> {
  const xml = await readDocxPartText(docx, partName);
  const parser = new DOMParser();
  const parsed = parser.parseFromString(xml, 'application/xml');
  return parsed.getElementsByTagName(tagName).length;
}
