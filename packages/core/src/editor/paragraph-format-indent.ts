/**
 * Write signed character-unit first-line indent (`w:firstLineChars` / `w:hangingChars`)
 * onto a `w:ind` attribute bag alongside an optional twip fallback (§17.3.1.12).
 */
export function writeFirstLineChars(
  authored: Readonly<Record<string, string>>,
  chars: number | null,
  fallbackTwips: number | null | undefined
): Record<string, string> {
  const attributes: Record<string, string> = { ...authored };
  delete attributes.firstLineChars;
  delete attributes.hangingChars;
  if (chars === null) {
    delete attributes.firstLine;
    delete attributes.hanging;
    return attributes;
  }
  if (chars === 0) {
    attributes.firstLine = '0';
    attributes.firstLineChars = '0';
    attributes.hanging = '0';
    attributes.hangingChars = '0';
    return attributes;
  }
  if (chars > 0) {
    attributes.firstLineChars = String(chars);
    attributes.hanging = '0';
    if (fallbackTwips !== undefined && fallbackTwips !== null) {
      attributes.firstLine = String(Math.max(0, fallbackTwips));
    } else {
      delete attributes.firstLine;
    }
  } else {
    attributes.hangingChars = String(-chars);
    attributes.firstLine = '0';
    if (fallbackTwips !== undefined && fallbackTwips !== null) {
      attributes.hanging = String(Math.max(0, -fallbackTwips));
    } else {
      delete attributes.hanging;
    }
  }
  return attributes;
}
