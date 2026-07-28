'use strict';

// Port of PptxToHtml/pptx/ThemeHelper.cs

const { firstChild, elements, attr } = require('../ooxml/dom');
const { slideLayoutPart, slideMasterPart, themePart } = require('../ooxml/parts');

const COLOR_SLOTS = [
  ['dk1', 'dk1'], ['lt1', 'lt1'], ['dk2', 'dk2'], ['lt2', 'lt2'],
  ['accent1', 'accent1'], ['accent2', 'accent2'], ['accent3', 'accent3'],
  ['accent4', 'accent4'], ['accent5', 'accent5'], ['accent6', 'accent6'],
  ['hlink', 'hlink'], ['folHlink', 'folHlink'],
];

const CLR_MAP_TARGETS = new Set([
  'dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3',
  'accent4', 'accent5', 'accent6', 'hlink', 'folHlink',
]);

const CLR_MAP_NAMES = ['tx1', 'bg1', 'tx2', 'bg2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

module.exports = {
  loadThemeColors(slidePart) {
    this._themeColors.clear();

    const masterPart = slideMasterPart(slideLayoutPart(slidePart));
    const theme = themePart(masterPart);
    const themeRoot = theme ? theme.root : null;
    const themeElements = firstChild(themeRoot, 'themeElements');
    const cs = firstChild(themeElements, 'clrScheme');
    if (cs == null) return;

    const add = (key, slotName) => {
      const el = firstChild(cs, slotName);
      if (el == null) return;
      const hex =
        attr(firstChild(el, 'srgbClr'), 'val') ?? attr(firstChild(el, 'sysClr'), 'lastClr');
      if (hex != null) this._themeColors.set(key, hex);
    };
    for (const [key, slot] of COLOR_SLOTS) add(key, slot);

    // Theme fonts
    const fontScheme = firstChild(themeElements, 'fontScheme');
    const majorLatin = firstChild(firstChild(fontScheme, 'majorFont'), 'latin');
    const minorLatin = firstChild(firstChild(fontScheme, 'minorFont'), 'latin');
    this._themeMajorFont = this.getStringValue(attr(majorLatin, 'typeface'));
    this._themeMinorFont = this.getStringValue(attr(minorLatin, 'typeface'));

    const masterRoot = masterPart ? masterPart.root : null;
    this._masterOtherStyle = firstChild(firstChild(masterRoot, 'txStyles'), 'otherStyle');

    // Theme line style widths
    this._themeLineWidths.length = 0;
    const fmtScheme = firstChild(themeElements, 'fmtScheme');
    const lnStyleLst = firstChild(fmtScheme, 'lnStyleLst');
    if (lnStyleLst != null) {
      for (const ln of elements(lnStyleLst, 'ln')) {
        const w = attr(ln, 'w');
        this._themeLineWidths.push(w != null ? Number(w) : 12700);
      }
    }

    // Color map from slide master
    this._colorMap.clear();
    const clrMap = firstChild(masterRoot, 'clrMap');
    if (clrMap != null) {
      for (const name of CLR_MAP_NAMES) {
        const val = attr(clrMap, name);
        if (val != null && CLR_MAP_TARGETS.has(val)) {
          this._colorMap.set(name, val);
        }
      }
    }
  },
};
