import JSZip from 'jszip';

// Fixed optional content parts beyond word/document.xml, in processing order.
// Headers and footers are discovered dynamically from the zip manifest because
// their numbering is not fixed (header1.xml, header2.xml, …).
// Note: headers/footers were an oversight in the original C# engine; comments
// support follows more-recent OOXML spec additions.
const FIXED_OPTIONAL_PARTS = [
  'word/footnotes.xml',
  'word/endnotes.xml',
  'word/comments.xml',
] as const;

const HEADER_FOOTER_RE = /^word\/(header|footer)\d+\.xml$/;

// Returns content parts in a stable processing order:
//   word/document.xml first, then fixed optional parts in declaration order,
//   then headers and footers sorted alphabetically.
export function collectContentParts(zip: JSZip): string[] {
  const parts: string[] = ['word/document.xml'];

  for (const partName of FIXED_OPTIONAL_PARTS) {
    if (zip.file(partName) !== null) {
      parts.push(partName);
    }
  }

  for (const partName of Object.keys(zip.files).sort()) {
    if (HEADER_FOOTER_RE.test(partName)) {
      parts.push(partName);
    }
  }

  return parts;
}
