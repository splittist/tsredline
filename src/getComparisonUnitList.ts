import { type ComparisonUnitAtom } from './createAtomList';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type ComparisonUnitGroupKind = 'paragraph' | 'table' | 'row' | 'cell' | 'textbox';

export interface ComparisonUnitWord {
  readonly kind: 'word';
  readonly sha1Hash: string;
  readonly atoms: readonly ComparisonUnitAtom[];
}

export interface ComparisonUnitGroup {
  readonly kind: ComparisonUnitGroupKind;
  readonly sha1Hash: string;
  readonly contents: readonly ComparisonUnit[];
}

export type ComparisonUnit = ComparisonUnitWord | ComparisonUnitGroup;

// ---------------------------------------------------------------------------
// SHA-1 helper (mirrors hashBlockLevelContent)
// ---------------------------------------------------------------------------

function bytesToHex(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let out = '';
  for (const b of bytes) out += b.toString(16).padStart(2, '0');
  return out;
}

async function sha1Hex(value: string): Promise<string> {
  const bytes = new TextEncoder().encode(value);
  const digest = await globalThis.crypto.subtle.digest('SHA-1', bytes);
  return bytesToHex(digest);
}

// ---------------------------------------------------------------------------
// groupAdjacent — groups consecutive items that share the same key
// ---------------------------------------------------------------------------

function groupAdjacent<T>(
  items: readonly T[],
  keyOf: (item: T) => string,
): Array<{ key: string; items: T[] }> {
  const result: Array<{ key: string; items: T[] }> = [];
  for (const item of items) {
    const key = keyOf(item);
    const last = result[result.length - 1];
    if (last !== undefined && last.key === key) {
      last.items.push(item);
    } else {
      result.push({ key, items: [item] });
    }
  }
  return result;
}

// ---------------------------------------------------------------------------
// Internal type — word unit with its ancestor key path
// ---------------------------------------------------------------------------

interface WithGroupingKey {
  readonly word: ComparisonUnitWord;
  readonly ancestorKeys: readonly string[];
}

// localName prefix → ComparisonUnitGroupKind
const PREFIX_TO_KIND: Readonly<Record<string, ComparisonUnitGroupKind>> = {
  p: 'paragraph',
  tbl: 'table',
  tr: 'row',
  tc: 'cell',
  txbxContent: 'textbox',
};

// ---------------------------------------------------------------------------
// Recursive hierarchy builder — mirrors C# GetHierarchicalComparisonUnits
// ---------------------------------------------------------------------------

async function getHierarchicalComparisonUnits(
  items: readonly WithGroupingKey[],
  level: number,
): Promise<readonly ComparisonUnit[]> {
  // Group adjacent items by the key at this level ("" when past the end).
  const grouped = groupAdjacent(items, (item) =>
    level < item.ancestorKeys.length ? item.ancestorKeys[level]! : '',
  );

  const result: ComparisonUnit[] = [];

  for (const { key, items: groupItems } of grouped) {
    if (key === '') {
      // Past the hierarchy depth — emit words directly.
      for (const { word } of groupItems) result.push(word);
    } else {
      const prefix = key.split(':')[0] ?? '';
      const kind: ComparisonUnitGroupKind = PREFIX_TO_KIND[prefix] ?? 'paragraph';
      const children = await getHierarchicalComparisonUnits(groupItems, level + 1);
      const sha1Hash = await sha1Hex(children.map((c) => c.sha1Hash).join(''));
      result.push({ kind, sha1Hash, contents: children });
    }
  }

  return result;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Groups a flat atom list into a hierarchy of ComparisonUnitWord and
 * ComparisonUnitGroup objects (paragraph / cell / row / table / textbox).
 *
 * Mirrors C# GetComparisonUnitList → GetHierarchicalComparisonUnits.
 * Each atom becomes one ComparisonUnitWord; words are then wrapped in
 * structural groups based on atom.ancestorKeys.
 */
export async function getComparisonUnitList(
  atoms: readonly ComparisonUnitAtom[],
): Promise<readonly ComparisonUnit[]> {
  // Build word units from atoms. Each atom is already word-level in our port,
  // so every atom maps to exactly one ComparisonUnitWord.
  const withKeys: WithGroupingKey[] = await Promise.all(
    atoms.map(async (atom) => {
      const sha1Hash = await sha1Hex(atom.kind + ':' + atom.text);
      const word: ComparisonUnitWord = { kind: 'word', sha1Hash, atoms: [atom] };
      return { word, ancestorKeys: atom.ancestorKeys };
    }),
  );

  return getHierarchicalComparisonUnits(withKeys, 0);
}
