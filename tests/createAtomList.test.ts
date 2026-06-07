import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';

import { assignUnids } from '../src/assignUnids';
import { createAtomList } from '../src/createAtomList';
import { loadFixture } from './helpers/fixtureLoader';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

async function makeDocx(documentXml: string): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'arraybuffer' });
}

// Build a synthetic DOCX, stamp it, and return the stamped buffer.
async function stampedDocx(documentXml: string): Promise<ArrayBuffer> {
  const raw = await makeDocx(documentXml);
  const { docx } = await assignUnids(raw);
  return docx;
}

// ---------------------------------------------------------------------------
// Word splitting
// ---------------------------------------------------------------------------

describe('createAtomList – word splitting', () => {
  it('splits a single run into one atom per word', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>hello world foo</w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    expect(atoms.map((a) => a.text)).toEqual(['hello', 'world', 'foo', '<w:pPr/>']);
    expect(atoms.slice(0, 3).every((a) => a.kind === 'word')).toBe(true);
    expect(atoms[3]!.kind).toBe('paragraph-mark');
  });

  it('handles leading/trailing whitespace in w:t correctly', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t xml:space="preserve">  alpha  beta  </w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    expect(atoms.filter((a) => a.kind === 'word').map((a) => a.text)).toEqual(['alpha', 'beta']);
  });

  it('concatenates words from consecutive runs in the same paragraph', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r><w:t>one</w:t></w:r>
          <w:r><w:t>two</w:t></w:r>
          <w:r><w:t>three</w:t></w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    expect(atoms.filter((a) => a.kind === 'word').map((a) => a.text)).toEqual(['one', 'two', 'three']);
  });

  it('emits only a paragraph-mark atom for a whitespace-only paragraph', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>   </w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    expect(atoms).toHaveLength(1);
    expect(atoms[0]!.kind).toBe('paragraph-mark');
  });
});

// ---------------------------------------------------------------------------
// Paragraph identity
// ---------------------------------------------------------------------------

describe('createAtomList – paraUnid', () => {
  it('assigns the correct paraUnid to atoms from each paragraph', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>alpha</w:t></w:r></w:p>
        <w:p><w:r><w:t>beta</w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const words = atoms.filter((a) => a.kind === 'word');
    expect(words).toHaveLength(2);
    const [a, b] = words as [typeof words[0], typeof words[0]];
    expect(a.paraUnid).toBeGreaterThan(0);
    expect(b.paraUnid).toBeGreaterThan(0);
    expect(a.paraUnid).not.toBe(b.paraUnid);
    expect(a.text).toBe('alpha');
    expect(b.text).toBe('beta');
  });

  it('groups all words from a multi-word paragraph under one paraUnid', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r><w:t>one two</w:t></w:r>
          <w:r><w:t>three</w:t></w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const words = atoms.filter((a) => a.kind === 'word');
    expect(words).toHaveLength(3);
    const unids = words.map((a) => a.paraUnid);
    expect(new Set(unids).size).toBe(1);
    expect(unids[0]).toBeGreaterThan(0);
  });
});

// ---------------------------------------------------------------------------
// Run and paragraph properties
// ---------------------------------------------------------------------------

describe('createAtomList – run and paragraph properties', () => {
  it('captures w:rPr XML on each word atom', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r>
            <w:rPr><w:b/></w:rPr>
            <w:t>bold</w:t>
          </w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const word = atoms.find((a) => a.kind === 'word');
    expect(word).toBeDefined();
    expect(word!.runPropsXml).toContain('w:b');
  });

  it('captures w:pPr XML on atoms from a styled paragraph', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
          <w:r><w:t>heading</w:t></w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const headingWord = atoms.find((a) => a.kind === 'word');
    expect(headingWord).toBeDefined();
    expect(headingWord!.paraPropsXml).toContain('Heading1');
  });

  it('propagates the same runPropsXml to all words from one run', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r>
            <w:rPr><w:i/></w:rPr>
            <w:t>italic words here</w:t>
          </w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const words = atoms.filter((a) => a.kind === 'word');
    expect(words).toHaveLength(3);
    const xmls = words.map((a) => a.runPropsXml);
    expect(new Set(xmls).size).toBe(1);
    expect(xmls[0]).toContain('w:i');
  });

  it('uses empty string for runPropsXml when the run has no w:rPr', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>plain</w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const word = atoms.find((a) => a.kind === 'word');
    expect(word).toBeDefined();
    expect(word!.runPropsXml).toBe('');
  });

  it('captures a positive runUnid for run-derived atoms', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r><w:t>alpha beta</w:t></w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const words = atoms.filter((a) => a.kind === 'word');
    expect(words).toHaveLength(2);
    expect(words.every((a) => a.runUnid > 0)).toBe(true);
    expect(new Set(words.map((a) => a.runUnid)).size).toBe(1);
  });

  it('uses runUnid=0 for paragraph-level atoms', async () => {
    const docx = await stampedDocx(`
      <w:document ${W} ${M}><w:body>
        <w:p>
          <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const paragraphMark = atoms.find((a) => a.kind === 'paragraph-mark');
    const math = atoms.find((a) => a.kind === 'math');
    expect(paragraphMark).toBeDefined();
    expect(math).toBeDefined();
    expect(paragraphMark!.runUnid).toBe(0);
    expect(math!.runUnid).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Non-text inline atoms
// ---------------------------------------------------------------------------

describe('createAtomList – non-text inline atoms', () => {
  it('emits a break atom for w:br inside a run', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r><w:t>before</w:t></w:r>
          <w:r><w:br/></w:r>
          <w:r><w:t>after</w:t></w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const kinds = atoms.filter((a) => a.kind !== 'paragraph-mark').map((a) => a.kind);
    expect(kinds).toContain('break');
    expect(kinds).toEqual(['word', 'break', 'word']);
  });

  it('emits a math atom for m:oMath at paragraph level', async () => {
    const docx = await stampedDocx(`
      <w:document ${W} ${M}><w:body>
        <w:p>
          <w:r><w:t>see</w:t></w:r>
          <m:oMath><m:r><m:t>x+1</m:t></m:r></m:oMath>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    expect(atoms.some((a) => a.kind === 'math')).toBe(true);
    const mathAtom = atoms.find((a) => a.kind === 'math')!;
    expect(mathAtom.text).toContain('oMath');
    expect(mathAtom.runPropsXml).toBe('');
  });

  it('emits an image atom for w:drawing inside a run', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"/></w:drawing></w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const imageAtoms = atoms.filter((a) => a.kind === 'image');
    expect(imageAtoms).toHaveLength(1);
    expect(imageAtoms[0]!.text).toContain('drawing');
  });

  it('two identical break atoms have the same text key', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r><w:br/></w:r>
          <w:r><w:br/></w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const breaks = atoms.filter((a) => a.kind === 'break');
    expect(breaks).toHaveLength(2);
    expect(breaks[0]!.text).toBe(breaks[1]!.text);
  });
});

// ---------------------------------------------------------------------------
// Run containers (hyperlinks, revision wrappers, etc.)
// ---------------------------------------------------------------------------

describe('createAtomList – run containers', () => {
  it('extracts words from runs inside w:hyperlink', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:hyperlink w:id="rId1">
            <w:r><w:t>click here</w:t></w:r>
          </w:hyperlink>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    expect(atoms.filter((a) => a.kind === 'word').map((a) => a.text)).toEqual(['click', 'here']);
  });

  it('respects revisionMode when processing revision markup', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r><w:t>kept</w:t></w:r>
          <w:ins w:id="1"><w:r><w:t>inserted</w:t></w:r></w:ins>
          <w:del w:id="2"><w:r><w:delText>deleted</w:delText></w:r></w:del>
        </w:p>
      </w:body></w:document>
    `);

    const accepted = await createAtomList(docx, { revisionMode: 'accept' });
    expect(accepted.atoms.filter((a) => a.kind === 'word').map((a) => a.text)).toEqual(['kept', 'inserted']);

    const rejected = await createAtomList(docx, { revisionMode: 'reject' });
    expect(rejected.atoms.filter((a) => a.kind === 'word').map((a) => a.text)).toEqual(['kept', 'deleted']);
  });
});

// ---------------------------------------------------------------------------
// atomCount invariant
// ---------------------------------------------------------------------------

describe('createAtomList – atomCount', () => {
  it('atomCount always equals atoms.length', async () => {
    const docx = await loadFixture('WC', 'WC002-Unmodified.docx');
    const { docx: stamped } = await assignUnids(docx);
    const result = await createAtomList(stamped);
    expect(result.atomCount).toBe(result.atoms.length);
  });
});

// ---------------------------------------------------------------------------
// Multi-part processing
// ---------------------------------------------------------------------------

describe('createAtomList – multiple content parts', () => {
  it('processes word/document.xml for a plain fixture', async () => {
    const docx = await loadFixture('CA', 'CA001-Plain.docx');
    const { docx: stamped } = await assignUnids(docx);
    const result = await createAtomList(stamped);

    expect(result.partNames).toContain('word/document.xml');
    expect(result.atomCount).toBeGreaterThan(0);
  });

  it('processes both document and footnotes parts when footnotes are present', async () => {
    const docx = await loadFixture('WC', 'WC020-FootNote-Before.docx');
    const { docx: stamped } = await assignUnids(docx);
    const result = await createAtomList(stamped);

    expect(result.partNames).toContain('word/document.xml');
    expect(result.partNames).toContain('word/footnotes.xml');
  });

  it('word atom count matches exactly for synthetic content with known words', async () => {
    // Use a synthetic document so the expected count is certain.
    // preprocessDocx concatenates adjacent w:t without separators before splitting,
    // which gives different counts when words are split across runs.
    // createAtomList splits within each w:t independently, which is the correct
    // behaviour for word-level comparison units.
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>one two three</w:t></w:r></w:p>
        <w:p><w:r><w:t>four five</w:t></w:r></w:p>
        <w:p><w:r><w:t>six</w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const result = await createAtomList(docx);
    const wordAtomCount = result.atoms.filter((a) => a.kind === 'word').length;
    expect(wordAtomCount).toBe(6);
  });

  it('produces the same atom sequence on repeated calls (deterministic)', async () => {
    const docx = await loadFixture('WC', 'WC004-Large.docx');
    const { docx: stamped } = await assignUnids(docx);

    const first = await createAtomList(stamped);
    const second = await createAtomList(stamped);

    expect(first.atoms.map((a) => a.text)).toEqual(second.atoms.map((a) => a.text));
  });
});

// ---------------------------------------------------------------------------
// Table structure
// ---------------------------------------------------------------------------

describe('createAtomList – table content', () => {
  it('visits paragraphs inside table cells', async () => {
    const docx = await loadFixture('WC', 'WC006-Table.docx');
    const { docx: stamped } = await assignUnids(docx);
    const result = await createAtomList(stamped);

    // Table fixture has text in cells; atoms must be produced.
    expect(result.atomCount).toBeGreaterThan(0);

    // All produced atoms must have a valid paraUnid (stamped during assignUnids).
    const wordAtoms = result.atoms.filter((a) => a.kind === 'word');
    expect(wordAtoms.every((a) => a.paraUnid > 0)).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// ancestorKeys propagation
// ---------------------------------------------------------------------------

describe('createAtomList – ancestorKeys', () => {
  it('standalone paragraph atom has ancestorKeys ending with p:N', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>hello</w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const word = atoms.find((a) => a.kind === 'word');
    expect(word).toBeDefined();
    const keys = word!.ancestorKeys;
    expect(keys).toHaveLength(1);
    expect(keys[0]).toMatch(/^p:\d+$/);
  });

  it('table-cell paragraph atom has ancestorKeys [tbl:X, tr:Y, tc:Z, p:N]', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:tbl>
          <w:tr>
            <w:tc>
              <w:p><w:r><w:t>cell</w:t></w:r></w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const word = atoms.find((a) => a.kind === 'word');
    expect(word).toBeDefined();
    const keys = word!.ancestorKeys;
    expect(keys).toHaveLength(4);
    expect(keys[0]).toMatch(/^tbl:\d+$/);
    expect(keys[1]).toMatch(/^tr:\d+$/);
    expect(keys[2]).toMatch(/^tc:\d+$/);
    expect(keys[3]).toMatch(/^p:\d+$/);
  });

  it('all atoms in a paragraph share the same ancestorKeys', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>one two three</w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const words = atoms.filter((a) => a.kind === 'word');
    const firstKeys = words[0]!.ancestorKeys;
    for (const atom of words) {
      expect(atom.ancestorKeys).toEqual(firstKeys);
    }
  });
});

// ---------------------------------------------------------------------------
// Paragraph-mark atoms
// ---------------------------------------------------------------------------

describe('createAtomList – paragraph-mark atoms', () => {
  it('emits one paragraph-mark atom per paragraph', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>one</w:t></w:r></w:p>
        <w:p><w:r><w:t>two</w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const marks = atoms.filter((a) => a.kind === 'paragraph-mark');
    expect(marks).toHaveLength(2);
    expect(marks.every((m) => m.runPropsXml === '')).toBe(true);
  });

  it('paragraph-mark atom carries serialized pPr when present', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
          <w:r><w:t>heading</w:t></w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const mark = atoms.find((a) => a.kind === 'paragraph-mark');
    expect(mark).toBeDefined();
    expect(mark!.text).toContain('w:pPr');
    expect(mark!.text).toContain('Heading1');
  });
});
