import { describe, expect, it } from 'vitest';

import { compareDocx } from '../src';

function toArrayBuffer(value: string): ArrayBuffer {
  return new TextEncoder().encode(value).buffer;
}

describe('compareDocx', () => {
  it('reports identical inputs as equal', async () => {
    const result = await compareDocx(
      toArrayBuffer('same document'),
      toArrayBuffer('same document'),
    );

    expect(result.equal).toBe(true);
    expect(result.changes).toHaveLength(0);
    expect(result.metadata.identicalBinary).toBe(true);
    expect(result.notices[0]?.code).toBe('SKELETON_ENGINE');
  });

  it('reports differing inputs as a placeholder replacement', async () => {
    const result = await compareDocx(
      toArrayBuffer('baseline document'),
      toArrayBuffer('candidate document'),
    );

    expect(result.equal).toBe(false);
    expect(result.changes).toEqual([
      {
        kind: 'replace',
        path: 'docx-package',
        before: '17 bytes',
        after: '18 bytes',
      },
    ]);
    expect(result.metadata.identicalBinary).toBe(false);
  });
});
