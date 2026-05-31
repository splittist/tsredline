export function arrayBuffersEqual(
  left: ArrayBuffer,
  right: ArrayBuffer,
): boolean {
  if (left.byteLength !== right.byteLength) {
    return false;
  }

  const leftView = new Uint8Array(left);
  const rightView = new Uint8Array(right);

  for (let index = 0; index < leftView.length; index += 1) {
    if (leftView[index] !== rightView[index]) {
      return false;
    }
  }

  return true;
}
