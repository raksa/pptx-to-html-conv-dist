'use strict';

// Port of PptxToHtml/pptx/ColorHelper.cs

const { firstChild, attr, childElements } = require('../ooxml/dom');
const { hex2, bankRound } = require('../util/format');
const { CIMap } = require('../util/collections');

// Map common OOXML preset color names to hex (StringComparer.OrdinalIgnoreCase).
const PRESET_COLORS = new CIMap([
  ['white', '#FFFFFF'], ['black', '#000000'], ['red', '#FF0000'],
  ['green', '#008000'], ['blue', '#0000FF'], ['yellow', '#FFFF00'],
  ['cyan', '#00FFFF'], ['magenta', '#FF00FF'], ['gray', '#808080'],
  ['darkGray', '#A9A9A9'], ['lightGray', '#D3D3D3'], ['darkRed', '#8B0000'],
  ['darkGreen', '#006400'], ['darkBlue', '#00008B'], ['navy', '#000080'],
  ['orange', '#FFA500'], ['pink', '#FFC0CB'], ['purple', '#800080'],
  ['maroon', '#800000'], ['olive', '#808000'], ['silver', '#C0C0C0'],
  ['aqua', '#00FFFF'], ['lime', '#00FF00'], ['teal', '#008080'],
  ['fuchsia', '#FF00FF'], ['brown', '#A52A2A'],
]);

const KNOWN_SCHEME_NAMES = new Set([
  'tx1', 'bg1', 'tx2', 'bg2', 'dk1', 'lt1', 'dk2', 'lt2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
]);

module.exports = {
  PRESET_COLORS,

  getColorFromFill(fill) {
    const srgb = firstChild(fill, 'srgbClr');
    if (srgb && attr(srgb, 'val') != null) {
      let hex = attr(srgb, 'val');
      hex = this.applyColorTransforms(hex, childElements(srgb));
      return `#${hex}`;
    }

    const schemeColor = firstChild(fill, 'schemeClr');
    if (schemeColor && attr(schemeColor, 'val') != null) {
      let hex = this.resolveSchemeColor(attr(schemeColor, 'val'));
      hex = this.applyColorTransforms(hex, childElements(schemeColor));
      return `#${hex}`;
    }

    const presetClr = firstChild(fill, 'prstClr');
    if (presetClr && attr(presetClr, 'val') != null) {
      const name = attr(presetClr, 'val');
      if (PRESET_COLORS.has(name)) return PRESET_COLORS.get(name);
    }

    return '';
  },

  getColorFromHighlight(highlight) {
    const srgb = firstChild(highlight, 'srgbClr');
    if (srgb && attr(srgb, 'val') != null) {
      let hex = attr(srgb, 'val');
      hex = this.applyColorTransforms(hex, childElements(srgb));
      return `#${hex}`;
    }

    const schemeColor = firstChild(highlight, 'schemeClr');
    if (schemeColor && attr(schemeColor, 'val') != null) {
      let hex = this.resolveSchemeColor(attr(schemeColor, 'val'));
      hex = this.applyColorTransforms(hex, childElements(schemeColor));
      return `#${hex}`;
    }

    const presetClr = firstChild(highlight, 'prstClr');
    if (presetClr && attr(presetClr, 'val') != null) {
      const name = attr(presetClr, 'val');
      if (PRESET_COLORS.has(name)) return PRESET_COLORS.get(name);
    }

    return '';
  },

  getColorFromColorType(stop) {
    const srgb = firstChild(stop, 'srgbClr');
    if (srgb && attr(srgb, 'val') != null) {
      let hex = attr(srgb, 'val');
      hex = this.applyColorTransforms(hex, childElements(srgb));
      return `#${hex}`;
    }

    const schemeColor = firstChild(stop, 'schemeClr');
    if (schemeColor && attr(schemeColor, 'val') != null) {
      let hex = this.resolveSchemeColor(attr(schemeColor, 'val'));
      hex = this.applyColorTransforms(hex, childElements(schemeColor));
      return `#${hex}`;
    }

    return '#000000';
  },

  /** Resolve a scheme color token (e.g. "tx1", "accent1") to a hex color (no #). */
  resolveSchemeColor(val) {
    if (val == null || val === '') return '000000';

    const schemeName = KNOWN_SCHEME_NAMES.has(val) ? val : 'dk1';

    let key;
    if (this._colorMap.has(schemeName)) {
      key = this._colorMap.get(schemeName);
    } else if (schemeName === 'tx1') key = 'dk1';
    else if (schemeName === 'bg1') key = 'lt1';
    else if (schemeName === 'tx2') key = 'dk2';
    else if (schemeName === 'bg2') key = 'lt2';
    else key = schemeName;

    return this._themeColors.has(key) ? this._themeColors.get(key) : '000000';
  },

  /** Apply tint/shade/lumMod/lumOff/alpha transforms to a hex color. */
  applyColorTransforms(hex, children) {
    if (hex.length < 6) return hex;

    let r = parseInt(hex.substr(0, 2), 16);
    let g = parseInt(hex.substr(2, 2), 16);
    let b = parseInt(hex.substr(4, 2), 16);
    let alphaPercent = 100;

    const clamp = (v) => Math.max(0, Math.min(255, v));

    for (const child of children) {
      const val = attr(child, 'val');
      if (val == null) continue;
      const v = parseInt(val, 10);
      if (child.localName === 'tint') {
        const t = v / 100000.0;
        r = clamp(bankRound(r + (255 - r) * t));
        g = clamp(bankRound(g + (255 - g) * t));
        b = clamp(bankRound(b + (255 - b) * t));
      } else if (child.localName === 'shade') {
        const s = v / 100000.0;
        r = clamp(bankRound(r * s));
        g = clamp(bankRound(g * s));
        b = clamp(bankRound(b * s));
      } else if (child.localName === 'lumMod') {
        let [h, s, l] = this.rgbToHsl(r, g, b);
        l = Math.max(0.0, Math.min(1.0, l * (v / 100000.0)));
        [r, g, b] = this.hslToRgb(h, s, l);
      } else if (child.localName === 'lumOff') {
        let [h, s, l] = this.rgbToHsl(r, g, b);
        l = Math.max(0.0, Math.min(1.0, l + v / 100000.0));
        [r, g, b] = this.hslToRgb(h, s, l);
      } else if (child.localName === 'alpha') {
        alphaPercent = Math.trunc(v / 1000);
      }
    }

    let result = `${hex2(clamp(r))}${hex2(clamp(g))}${hex2(clamp(b))}`;
    if (alphaPercent < 100) {
      const a = bankRound((255.0 * Math.max(0, Math.min(100, alphaPercent))) / 100.0);
      result += hex2(Math.max(0, Math.min(255, a)));
    }
    return result;
  },

  rgbToHsl(r, g, b) {
    const rf = r / 255.0;
    const gf = g / 255.0;
    const bf = b / 255.0;
    const max = Math.max(rf, gf, bf);
    const min = Math.min(rf, gf, bf);
    let h = 0;
    const l = (max + min) / 2.0;
    const d = max - min;
    let s = 0;
    if (d > 0) {
      s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
      if (max === rf) h = ((gf - bf) / d + (gf < bf ? 6 : 0)) / 6.0;
      else if (max === gf) h = ((bf - rf) / d + 2) / 6.0;
      else h = ((rf - gf) / d + 4) / 6.0;
    }
    return [h, s, l];
  },

  hslToRgb(h, s, l) {
    const hueToRgb = (p, q, t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1.0 / 6.0) return p + (q - p) * 6 * t;
      if (t < 1.0 / 2.0) return q;
      if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6;
      return p;
    };

    let r, g, b;
    if (s === 0) {
      r = g = b = l;
    } else {
      const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
      const p = 2 * l - q;
      r = hueToRgb(p, q, h + 1.0 / 3.0);
      g = hueToRgb(p, q, h);
      b = hueToRgb(p, q, h - 1.0 / 3.0);
    }

    const clamp = (v) => Math.max(0, Math.min(255, bankRound(v * 255)));
    return [clamp(r), clamp(g), clamp(b)];
  },

  getThemeTextColor() {
    const key = this._colorMap.has('tx1') ? this._colorMap.get('tx1') : 'dk1';
    const hex = this._themeColors.has(key) ? this._themeColors.get(key) : '000000';
    return `#${hex}`;
  },

  resolveColorElement(colorContainer) {
    if (colorContainer == null) return '000000';

    const srgb = firstChild(colorContainer, 'srgbClr');
    if (srgb && attr(srgb, 'val') != null) {
      const hex = attr(srgb, 'val');
      return this.applyColorTransforms(hex, childElements(srgb));
    }

    const scheme = firstChild(colorContainer, 'schemeClr');
    if (scheme && attr(scheme, 'val') != null) {
      const hex = this.resolveSchemeColor(attr(scheme, 'val'));
      return this.applyColorTransforms(hex, childElements(scheme));
    }

    return '000000';
  },
};
