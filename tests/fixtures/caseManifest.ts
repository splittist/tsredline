export interface FixtureCase {
  readonly id: string;
  readonly baseline: { readonly group: 'WC' | 'CA'; readonly file: string };
  readonly candidate: { readonly group: 'WC' | 'CA'; readonly file: string };
  readonly category:
    | 'basic-text'
    | 'paragraph-boundary'
    | 'table'
    | 'move'
    | 'format'
    | 'footnote-endnote'
    | 'image-math';
  readonly expectedBehavior: string;
}

// Initial slice of C# WCB cases from reference/WmlComparerTests.cs.
export const CHARACTERIZATION_CASES: readonly FixtureCase[] = [
  {
    id: 'WCB-1000',
    baseline: { group: 'CA', file: 'CA001-Plain.docx' },
    candidate: { group: 'CA', file: 'CA001-Plain-Mod.docx' },
    category: 'basic-text',
    expectedBehavior: 'Should detect simple text edits in a plain document.',
  },
  {
    id: 'WCB-1010',
    baseline: { group: 'WC', file: 'WC001-Digits.docx' },
    candidate: { group: 'WC', file: 'WC001-Digits-Mod.docx' },
    category: 'basic-text',
    expectedBehavior: 'Should detect digit/text edits within paragraph content.',
  },
  {
    id: 'WCB-1020',
    baseline: { group: 'WC', file: 'WC001-Digits.docx' },
    candidate: { group: 'WC', file: 'WC001-Digits-Deleted-Paragraph.docx' },
    category: 'paragraph-boundary',
    expectedBehavior: 'Should mark paragraph deletion with correct paragraph mark behavior.',
  },
  {
    id: 'WCB-1070',
    baseline: { group: 'WC', file: 'WC002-Unmodified.docx' },
    candidate: { group: 'WC', file: 'WC002-InsertAtBeginning.docx' },
    category: 'paragraph-boundary',
    expectedBehavior: 'Should detect insertion at beginning without cascading unrelated diffs.',
  },
  {
    id: 'WCB-1110',
    baseline: { group: 'WC', file: 'WC002-Unmodified.docx' },
    candidate: { group: 'WC', file: 'WC002-InsertInMiddle.docx' },
    category: 'basic-text',
    expectedBehavior: 'Should detect middle insertion with stable surrounding matches.',
  },
  {
    id: 'WCB-1140',
    baseline: { group: 'WC', file: 'WC006-Table.docx' },
    candidate: { group: 'WC', file: 'WC006-Table-Delete-Row.docx' },
    category: 'table',
    expectedBehavior: 'Should detect deleted table row using table-aware correlation.',
  },
  {
    id: 'WCB-1160',
    baseline: { group: 'WC', file: 'WC006-Table.docx' },
    candidate: { group: 'WC', file: 'WC006-Table-Delete-Contests-of-Row.docx' },
    category: 'table',
    expectedBehavior: 'Should detect cell-content deletion while preserving row structure.',
  },
  {
    id: 'WCB-1190',
    baseline: { group: 'WC', file: 'WC007-Unmodified.docx' },
    candidate: { group: 'WC', file: 'WC007-Moved-into-Table.docx' },
    category: 'move',
    expectedBehavior: 'Should identify moved content instead of only delete/insert when move detection is enabled.',
  },
  {
    id: 'WCB-1230',
    baseline: { group: 'WC', file: 'WC012-Math-Before.docx' },
    candidate: { group: 'WC', file: 'WC012-Math-After.docx' },
    category: 'image-math',
    expectedBehavior: 'Should safely process and diff math content in WordprocessingML.',
  },
  {
    id: 'WCB-1240',
    baseline: { group: 'WC', file: 'WC013-Image-Before.docx' },
    candidate: { group: 'WC', file: 'WC013-Image-After.docx' },
    category: 'image-math',
    expectedBehavior: 'Should process image-related changes without corrupting document structure.',
  },
  {
    id: 'WCB-1380',
    baseline: { group: 'WC', file: 'WC020-FootNote-Before.docx' },
    candidate: { group: 'WC', file: 'WC020-FootNote-After-1.docx' },
    category: 'footnote-endnote',
    expectedBehavior: 'Should preserve and reconcile footnote references and IDs correctly.',
  },
  {
    id: 'WCB-1640',
    baseline: { group: 'WC', file: 'WC034-Endnotes-Before.docx' },
    candidate: { group: 'WC', file: 'WC034-Endnotes-After1.docx' },
    category: 'footnote-endnote',
    expectedBehavior: 'Should preserve and reconcile endnote references and IDs correctly.',
  },
] as const;
