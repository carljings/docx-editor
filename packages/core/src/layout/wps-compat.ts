// Opt-in WPS China Edition layout compatibility policy.
// Disabled by default so MS Word parity and upstream layout tests remain unchanged.

/**
 * Optional layout overrides for aligning with WPS China Edition official-document pagination.
 *
 * @public
 */
export interface WpsCompatOptions {
  /**
   * Enable CJK punctuation compression (`w:characterSpacingControl="compressPunctuation"`) and
   * line-end punctuation hanging on left-aligned (`w:jc="left"`) paragraphs.
   *
   * MS Word modern compatibility mode applies optical punctuation fitting only to justified
   * (`w:jc="both"`) paragraphs, whereas WPS China Edition compresses CJK punctuation on
   * left-aligned paragraphs as well.
   */
  readonly leftAlignedPunctuationCompression?: boolean;
}

let activeWpsCompat: Readonly<Required<WpsCompatOptions>> = Object.freeze({
  leftAlignedPunctuationCompression: false,
});

/**
 * Configure optional WPS China Edition layout compatibility overrides for the current runtime.
 *
 * @public
 */
export function setWpsCompatMode(options: WpsCompatOptions | boolean): void {
  if (typeof options === 'boolean') {
    activeWpsCompat = Object.freeze({
      leftAlignedPunctuationCompression: options,
    });
    return;
  }
  activeWpsCompat = Object.freeze({
    leftAlignedPunctuationCompression: Boolean(options.leftAlignedPunctuationCompression),
  });
}

/**
 * Read the active WPS China Edition layout compatibility policy.
 *
 * @public
 */
export function getWpsCompatMode(): Readonly<Required<WpsCompatOptions>> {
  return activeWpsCompat;
}
