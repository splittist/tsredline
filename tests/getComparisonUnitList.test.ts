import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';

import { assignUnids } from '../src/assignUnids';
import { createAtomList } from '../src/createAtomList';
import {
  getComparisonUnitList,
  type ComparisonUnit,
  type ComparisonUnitGroup,
  type ComparisonUnitWord,
} from '../src/getComparisonUnitList';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function makeDocx(documentXml: string): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function stampedDocx(documentXml: string): Promise<ArrayBuffer> {
  const raw = await makeDocx(documentXml);
  const { docx } = await assignUnids(raw);
  return docx;
}

async function atomsAndUnits(documentXml: string) {
  const docx = await stampedDocx(documentXml);
  const { atoms } = await createAtomList(docx);
  const units = await getComparisonUnitList(atoms);
  return { atoms, units };
}

function asGroup(unit: ComparisonUnit): ComparisonUnitGroup {
  if (unit.kind === 'word') throw new Error('Expected group, got word');
  return unit as ComparisonUnitGroup;
}

function asWord(unit: ComparisonUnit): ComparisonUnitWord {
  if (unit.kind !== 'word') throw new Error('Expected word, got group');
  return unit as ComparisonUnitWord;
}

// ---------------------------------------------------------------------------
// Single paragraph
// ---------------------------------------------------------------------------

describe('getComparisonUnitList – single paragraph', () => {
  it('wraps words in a paragraph group', async () => {
    const { units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>hello world</w:t></w:r></w:p>
      </w:body></w:document>
    `);

    expect(units).toHaveLength(1);
    const para = asGroup(units[0]!);
    expect(para.kind).toBe('paragraph');
    expect(para.contents).toHaveLength(2);
    expect(asWord(para.contents[0]!).atoms[0]!.text).toBe('hello');
    expect(asWord(para.contents[1]!).atoms[0]!.text).toBe('world');
  });

  it('every word unit wraps exactly one atom', async () => {
    const { atoms, units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>one two three</w:t></w:r></w:p>
      </w:body></w:document>
    `);

    const para = asGroup(units[0]!);
    expect(para.contents).toHaveLength(atoms.length);
    for (const cu of para.contents) {
      const w = asWord(cu);
      expect(w.atoms).toHaveLength(1);
    }
  });

  it('sha1Hash is a 40-char hex string on word units', async () => {
    const { units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>test</w:t></w:r></w:p>
      </w:body></w:document>
    `);

    const para = asGroup(units[0]!);
    const word = asWord(para.contents[0]!);
    expect(word.sha1Hash).toMatch(/^[0-9a-f]{40}$/);
  });

  it('sha1Hash on paragraph group is also a 40-char hex string', async () => {
    const { units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>test</w:t></w:r></w:p>
      </w:body></w:document>
    `);

    const para = asGroup(units[0]!);
    expect(para.sha1Hash).toMatch(/^[0-9a-f]{40}$/);
  });
});

// ---------------------------------------------------------------------------
// Multiple paragraphs
// ---------------------------------------------------------------------------

describe('getComparisonUnitList – multiple paragraphs', () => {
  it('produces one paragraph group per paragraph', async () => {
    const { units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>first</w:t></w:r></w:p>
        <w:p><w:r><w:t>second</w:t></w:r></w:p>
        <w:p><w:r><w:t>third</w:t></w:r></w:p>
      </w:body></w:document>
    `);

    expect(units).toHaveLength(3);
    for (const unit of units) expect(asGroup(unit).kind).toBe('paragraph');
  });

  it('paragraph groups have different sha1Hashes for different content', async () => {
    const { units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>alpha</w:t></w:r></w:p>
        <w:p><w:r><w:t>beta</w:t></w:r></w:p>
      </w:body></w:document>
    `);

    const [g1, g2] = units.map(asGroup);
    expect(g1!.sha1Hash).not.toBe(g2!.sha1Hash);
  });

  it('identical paragraphs have the same sha1Hash', async () => {
    const { units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>same</w:t></w:r></w:p>
        <w:p><w:r><w:t>same</w:t></w:r></w:p>
      </w:body></w:document>
    `);

    const [g1, g2] = units.map(asGroup);
    expect(g1!.sha1Hash).toBe(g2!.sha1Hash);
  });
});

// ---------------------------------------------------------------------------
// Table structure
// ---------------------------------------------------------------------------

describe('getComparisonUnitList – table hierarchy', () => {
  it('wraps table content in table > row > cell > paragraph groups', async () => {
    const { units } = await atomsAndUnits(`
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

    expect(units).toHaveLength(1);
    const table = asGroup(units[0]!);
    expect(table.kind).toBe('table');

    const row = asGroup(table.contents[0]!);
    expect(row.kind).toBe('row');

    const cell = asGroup(row.contents[0]!);
    expect(cell.kind).toBe('cell');

    const para = asGroup(cell.contents[0]!);
    expect(para.kind).toBe('paragraph');

    const word = asWord(para.contents[0]!);
    expect(word.atoms[0]!.text).toBe('cell');
  });

  it('produces one row group per table row', async () => {
    const { units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:tbl>
          <w:tr>
            <w:tc><w:p><w:r><w:t>r1c1</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>r1c2</w:t></w:r></w:p></w:tc>
          </w:tr>
          <w:tr>
            <w:tc><w:p><w:r><w:t>r2c1</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>r2c2</w:t></w:r></w:p></w:tc>
          </w:tr>
        </w:tbl>
      </w:body></w:document>
    `);

    const table = asGroup(units[0]!);
    expect(table.kind).toBe('table');
    expect(table.contents).toHaveLength(2);
    expect(asGroup(table.contents[0]!).kind).toBe('row');
    expect(asGroup(table.contents[1]!).kind).toBe('row');
  });

  it('produces one cell group per table cell within a row', async () => {
    const { units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:tbl>
          <w:tr>
            <w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc>
          </w:tr>
        </w:tbl>
      </w:body></w:document>
    `);

    const table = asGroup(units[0]!);
    const row = asGroup(table.contents[0]!);
    expect(row.contents).toHaveLength(3);
    for (const cu of row.contents) expect(asGroup(cu).kind).toBe('cell');
  });

  it('mixes standalone paragraphs and tables at the top level', async () => {
    const { units } = await atomsAndUnits(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>before</w:t></w:r></w:p>
        <w:tbl>
          <w:tr>
            <w:tc><w:p><w:r><w:t>in table</w:t></w:r></w:p></w:tc>
          </w:tr>
        </w:tbl>
        <w:p><w:r><w:t>after</w:t></w:r></w:p>
      </w:body></w:document>
    `);

    expect(units).toHaveLength(3);
    expect(asGroup(units[0]!).kind).toBe('paragraph');
    expect(asGroup(units[1]!).kind).toBe('table');
    expect(asGroup(units[2]!).kind).toBe('paragraph');
  });
});

// ---------------------------------------------------------------------------
// Determinism
// ---------------------------------------------------------------------------

describe('getComparisonUnitList – determinism', () => {
  it('produces identical hashes on repeated calls', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p><w:r><w:t>deterministic</w:t></w:r></w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);

    const run1 = await getComparisonUnitList(atoms);
    const run2 = await getComparisonUnitList(atoms);

    expect(asGroup(run1[0]!).sha1Hash).toBe(asGroup(run2[0]!).sha1Hash);
  });

  it('word hash differs for different texts', async () => {
    const docx = await stampedDocx(`
      <w:document ${W}><w:body>
        <w:p>
          <w:r><w:t>apple</w:t></w:r>
          <w:r><w:t>banana</w:t></w:r>
        </w:p>
      </w:body></w:document>
    `);
    const { atoms } = await createAtomList(docx);
    const units = await getComparisonUnitList(atoms);

    const para = asGroup(units[0]!);
    const w1 = asWord(para.contents[0]!);
    const w2 = asWord(para.contents[1]!);
    expect(w1.sha1Hash).not.toBe(w2.sha1Hash);
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
    expect(atoms).toHaveLength(1);
    const keys = atoms[0]!.ancestorKeys;
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
    expect(atoms).toHaveLength(1);
    const keys = atoms[0]!.ancestorKeys;
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
    const firstKeys = atoms[0]!.ancestorKeys;
    for (const atom of atoms) {
      expect(atom.ancestorKeys).toEqual(firstKeys);
    }
  });
});
