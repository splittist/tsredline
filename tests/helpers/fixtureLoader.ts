import { readFile } from 'node:fs/promises';
import { join } from 'node:path';

const FIXTURE_ROOT = join(process.cwd(), 'tests', 'fixtures');

export function fixturePath(group: 'WC' | 'CA', fileName: string): string {
  return join(FIXTURE_ROOT, group, fileName);
}

export async function loadFixture(group: 'WC' | 'CA', fileName: string): Promise<ArrayBuffer> {
  const bytes = await readFile(fixturePath(group, fileName));
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
}
