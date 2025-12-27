export function idxOf(header, col) {
  const normalize = (s) =>
    String(s)
      .replace(/\u00A0/g, " ") // NBSP → space
      .trim()
      .toLowerCase();

  const target = normalize(col);

  const idx = header.findIndex(
    (h) => normalize(h) === target
  );

  if (idx === -1) {
    throw new Error(`Missing column: ${col}`);
  }

  return idx;
}
