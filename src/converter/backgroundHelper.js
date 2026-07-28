'use strict';

// Port of PptxToHtml/pptx/BackgroundHelper.cs

const { firstChild, elements, attr } = require('../ooxml/dom');
const { slideLayoutPart, slideMasterPart } = require('../ooxml/parts');
const { f0, f1 } = require('../util/format');

function backgroundOf(root) {
  return firstChild(firstChild(root, 'cSld'), 'bg');
}

module.exports = {
  // ── SVG Background ──────────────────────────────────────────────

  getSlideBackgroundSvg(slidePart, presentationPart, slideWidthPx, slideHeightPx, defsSb) {
    const bg = backgroundOf(slidePart.root);
    if (bg != null) {
      const result = this.extractBackgroundSvg(bg, slidePart, slideWidthPx, slideHeightPx, defsSb);
      if (result !== '') return result;
    }

    const layoutPart = slideLayoutPart(slidePart);
    if (layoutPart != null) {
      const layoutBg = backgroundOf(layoutPart.root);
      if (layoutBg != null) {
        const result = this.extractBackgroundSvg(layoutBg, layoutPart, slideWidthPx, slideHeightPx, defsSb);
        if (result !== '') return result;
      }

      const masterPart = slideMasterPart(layoutPart);
      if (masterPart != null) {
        const masterBg = backgroundOf(masterPart.root);
        if (masterBg != null) {
          const result = this.extractBackgroundSvg(masterBg, masterPart, slideWidthPx, slideHeightPx, defsSb);
          if (result !== '') return result;
        }
      }
    }

    return `<rect width="${f0(slideWidthPx)}" height="${f0(slideHeightPx)}" fill="#ffffff"/>`;
  },

  extractBackgroundSvg(bg, part, width, height, defsSb) {
    const bgPr = firstChild(bg, 'bgPr');
    if (bgPr == null) return '';

    const solidFill = firstChild(bgPr, 'solidFill');
    if (solidFill != null) {
      const color = this.getColorFromFill(solidFill);
      if (color !== '') return `<rect width="${f0(width)}" height="${f0(height)}" ${this.svgFillAttr(color)}/>`;
    }

    const gradFill = firstChild(bgPr, 'gradFill');
    if (gradFill != null) {
      const [fillAttr, defs] = this.getSvgGradientFill(gradFill);
      if (defs !== '') defsSb.appendLine(defs);
      return `<rect width="${f0(width)}" height="${f0(height)}" ${fillAttr}/>`;
    }

    const blipFill = firstChild(bgPr, 'blipFill');
    const imgPart = this.resolveImagePart(part, firstChild(blipFill, 'blip'));
    if (imgPart != null) {
      const imgPath = this.saveImageFile(imgPart);
      return `<image href="${imgPath}" width="${f0(width)}" height="${f0(height)}" preserveAspectRatio="none"/>`;
    }

    return '';
  },

  // ── CSS Background (kept for gradient text fills) ────────────────

  getSlideBackground(slidePart, presentationPart) {
    const bg = backgroundOf(slidePart.root);
    if (bg != null) {
      const result = this.extractBackground(bg, slidePart);
      if (result !== '') return result;
    }

    const layoutPart = slideLayoutPart(slidePart);
    if (layoutPart != null) {
      const layoutBg = backgroundOf(layoutPart.root);
      if (layoutBg != null) {
        const result = this.extractBackground(layoutBg, layoutPart);
        if (result !== '') return result;
      }

      const masterPart = slideMasterPart(layoutPart);
      if (masterPart != null) {
        const masterBg = backgroundOf(masterPart.root);
        if (masterBg != null) {
          const result = this.extractBackground(masterBg, masterPart);
          if (result !== '') return result;
        }
      }
    }

    return 'background:#ffffff;';
  },

  extractBackground(bg, part) {
    const bgPr = firstChild(bg, 'bgPr');
    if (bgPr == null) return '';

    const solidFill = firstChild(bgPr, 'solidFill');
    if (solidFill != null) {
      const color = this.getColorFromFill(solidFill);
      if (color !== '') return `background-color:${color};`;
    }

    const gradFill = firstChild(bgPr, 'gradFill');
    if (gradFill != null) {
      const gradCss = this.getGradientFillCss(gradFill);
      if (gradCss !== '') return gradCss;
    }

    const blipFill = firstChild(bgPr, 'blipFill');
    const imgPart = this.resolveImagePart(part, firstChild(blipFill, 'blip'));
    if (imgPart != null) {
      const imgPath = this.saveImageFile(imgPart);
      return `background-image:url('${imgPath}');background-size:100% 100%;`;
    }

    return '';
  },

  getGradientFillCss(gradFill) {
    const stops = elements(firstChild(gradFill, 'gsLst'), 'gs');
    if (stops.length < 2) return '';

    let cssDeg = 180;
    const linFill = firstChild(gradFill, 'lin');
    const ang = attr(linFill, 'ang');
    if (ang != null) {
      const ooxmlDeg = Number(ang) / 60000.0;
      cssDeg = ((ooxmlDeg + 90) % 360 + 360) % 360;
    }

    const parts = [];
    for (const stop of stops) {
      const color = this.getColorFromColorType(stop);
      const posPct = (attr(stop, 'pos') != null ? Number(attr(stop, 'pos')) : 0) / 1000.0;
      parts.push(`${color} ${f1(posPct)}%`);
    }

    return `background:linear-gradient(${f0(cssDeg)}deg,${parts.join(',')});`;
  },

  getPatternFillCss(patternFill) {
    const fg = '#' + this.resolveColorElement(firstChild(patternFill, 'fgClr'));
    const bg = '#' + this.resolveColorElement(firstChild(patternFill, 'bgClr'));
    const preset = attr(patternFill, 'prst') || '';

    const svgPattern = this.generatePatternSvg(preset, fg, bg);
    const encoded = Buffer.from(svgPattern, 'utf8').toString('base64');
    return `background-image:url('data:image/svg+xml;base64,${encoded}');background-repeat:repeat;background-size:8px 8px;`;
  },

  generatePatternSvg(preset, fg, bg) {
    let lines;
    switch (preset) {
      case 'dnDiag':
        lines = `<line x1='0' y1='0' x2='8' y2='8' stroke='${fg}' stroke-width='1'/>`;
        break;
      case 'upDiag':
        lines = `<line x1='8' y1='0' x2='0' y2='8' stroke='${fg}' stroke-width='1'/>`;
        break;
      case 'horz':
        lines = `<line x1='0' y1='4' x2='8' y2='4' stroke='${fg}' stroke-width='1'/>`;
        break;
      case 'vert':
        lines = `<line x1='4' y1='0' x2='4' y2='8' stroke='${fg}' stroke-width='1'/>`;
        break;
      case 'cross':
        lines =
          `<line x1='0' y1='4' x2='8' y2='4' stroke='${fg}' stroke-width='1'/>` +
          `<line x1='4' y1='0' x2='4' y2='8' stroke='${fg}' stroke-width='1'/>`;
        break;
      case 'diagCross':
        lines =
          `<line x1='0' y1='0' x2='8' y2='8' stroke='${fg}' stroke-width='1'/>` +
          `<line x1='8' y1='0' x2='0' y2='8' stroke='${fg}' stroke-width='1'/>`;
        break;
      case 'pct5':
        lines = `<rect x='0' y='0' width='1' height='1' fill='${fg}'/>`;
        break;
      case 'pct10':
        lines = `<rect x='0' y='0' width='1' height='1' fill='${fg}'/><rect x='4' y='4' width='1' height='1' fill='${fg}'/>`;
        break;
      case 'pct20':
        lines =
          `<rect x='0' y='0' width='1' height='1' fill='${fg}'/><rect x='4' y='0' width='1' height='1' fill='${fg}'/>` +
          `<rect x='2' y='4' width='1' height='1' fill='${fg}'/><rect x='6' y='4' width='1' height='1' fill='${fg}'/>`;
        break;
      case 'pct25':
        lines =
          `<rect x='0' y='0' width='1' height='1' fill='${fg}'/><rect x='4' y='0' width='1' height='1' fill='${fg}'/>` +
          `<rect x='2' y='2' width='1' height='1' fill='${fg}'/><rect x='6' y='2' width='1' height='1' fill='${fg}'/>` +
          `<rect x='0' y='4' width='1' height='1' fill='${fg}'/><rect x='4' y='4' width='1' height='1' fill='${fg}'/>` +
          `<rect x='2' y='6' width='1' height='1' fill='${fg}'/><rect x='6' y='6' width='1' height='1' fill='${fg}'/>`;
        break;
      case 'pct50':
        lines =
          `<rect x='0' y='0' width='1' height='1' fill='${fg}'/><rect x='2' y='0' width='1' height='1' fill='${fg}'/>` +
          `<rect x='4' y='0' width='1' height='1' fill='${fg}'/><rect x='6' y='0' width='1' height='1' fill='${fg}'/>` +
          `<rect x='1' y='1' width='1' height='1' fill='${fg}'/><rect x='3' y='1' width='1' height='1' fill='${fg}'/>` +
          `<rect x='5' y='1' width='1' height='1' fill='${fg}'/><rect x='7' y='1' width='1' height='1' fill='${fg}'/>` +
          `<rect x='0' y='2' width='1' height='1' fill='${fg}'/><rect x='2' y='2' width='1' height='1' fill='${fg}'/>` +
          `<rect x='4' y='2' width='1' height='1' fill='${fg}'/><rect x='6' y='2' width='1' height='1' fill='${fg}'/>` +
          `<rect x='1' y='3' width='1' height='1' fill='${fg}'/><rect x='3' y='3' width='1' height='1' fill='${fg}'/>` +
          `<rect x='5' y='3' width='1' height='1' fill='${fg}'/><rect x='7' y='3' width='1' height='1' fill='${fg}'/>` +
          `<rect x='0' y='4' width='1' height='1' fill='${fg}'/><rect x='2' y='4' width='1' height='1' fill='${fg}'/>` +
          `<rect x='4' y='4' width='1' height='1' fill='${fg}'/><rect x='6' y='4' width='1' height='1' fill='${fg}'/>` +
          `<rect x='1' y='5' width='1' height='1' fill='${fg}'/><rect x='3' y='5' width='1' height='1' fill='${fg}'/>` +
          `<rect x='5' y='5' width='1' height='1' fill='${fg}'/><rect x='7' y='5' width='1' height='1' fill='${fg}'/>` +
          `<rect x='0' y='6' width='1' height='1' fill='${fg}'/><rect x='2' y='6' width='1' height='1' fill='${fg}'/>` +
          `<rect x='4' y='6' width='1' height='1' fill='${fg}'/><rect x='6' y='6' width='1' height='1' fill='${fg}'/>` +
          `<rect x='1' y='7' width='1' height='1' fill='${fg}'/><rect x='3' y='7' width='1' height='1' fill='${fg}'/>` +
          `<rect x='5' y='7' width='1' height='1' fill='${fg}'/><rect x='7' y='7' width='1' height='1' fill='${fg}'/>`;
        break;
      case 'pct75':
        return (
          `<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'>` +
          `<rect width='8' height='8' fill='${fg}'/>` +
          `<rect x='0' y='0' width='1' height='1' fill='${bg}'/><rect x='4' y='4' width='1' height='1' fill='${bg}'/>` +
          `</svg>`
        );
      case 'pct90':
        return (
          `<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'>` +
          `<rect width='8' height='8' fill='${fg}'/>` +
          `<rect x='0' y='0' width='1' height='1' fill='${bg}'/>` +
          `</svg>`
        );
      case 'dkDnDiag':
        lines = `<line x1='0' y1='0' x2='8' y2='8' stroke='${fg}' stroke-width='2'/>`;
        break;
      case 'dkUpDiag':
        lines = `<line x1='8' y1='0' x2='0' y2='8' stroke='${fg}' stroke-width='2'/>`;
        break;
      case 'dkHorz':
        lines = `<line x1='0' y1='4' x2='8' y2='4' stroke='${fg}' stroke-width='2'/>`;
        break;
      case 'dkVert':
        lines = `<line x1='4' y1='0' x2='4' y2='8' stroke='${fg}' stroke-width='2'/>`;
        break;
      case 'wdDnDiag':
        lines = `<line x1='-1' y1='-1' x2='9' y2='9' stroke='${fg}' stroke-width='3'/>`;
        break;
      case 'wdUpDiag':
        lines = `<line x1='9' y1='-1' x2='-1' y2='9' stroke='${fg}' stroke-width='3'/>`;
        break;
      case 'ltHorz':
        lines = `<line x1='0' y1='4' x2='8' y2='4' stroke='${fg}' stroke-width='0.5'/>`;
        break;
      case 'ltVert':
        lines = `<line x1='4' y1='0' x2='4' y2='8' stroke='${fg}' stroke-width='0.5'/>`;
        break;
      case 'ltDnDiag':
        lines = `<line x1='0' y1='0' x2='8' y2='8' stroke='${fg}' stroke-width='0.5'/>`;
        break;
      case 'ltUpDiag':
        lines = `<line x1='8' y1='0' x2='0' y2='8' stroke='${fg}' stroke-width='0.5'/>`;
        break;
      default:
        lines = `<line x1='0' y1='0' x2='8' y2='8' stroke='${fg}' stroke-width='1'/>`;
        break;
    }
    return (
      `<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'>` +
      `<rect width='8' height='8' fill='${bg}'/>` +
      lines +
      `</svg>`
    );
  },
};
