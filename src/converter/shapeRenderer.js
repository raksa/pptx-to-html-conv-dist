'use strict';

// Port of PptxToHtml/pptx/ShapeRenderer.cs

const { firstChild, elements, descendants, attr, attrLong, childElements, innerText, NS } = require('../ooxml/dom');
const { REL_TYPE } = require('../ooxml/package');
const { f0, f1, f2, f3, f4 } = require('../util/format');
const { placeholderOf, textBodyOf, shapePropsOf, shapeTreeOf } = require('./elements');

const A = NS.A;
const P = NS.P;
const C = NS.C;

const numOr = (el, name, def = 0) => {
  const v = attr(el, name);
  if (v == null) return def;
  const n = Number(v);
  return Number.isFinite(n) ? n : def;
};

const ciIncludes = (hay, needle) => (hay || '').toLowerCase().includes(needle.toLowerCase());
const ciEq = (a, b) => (a || '').toLowerCase() === (b || '').toLowerCase();

function geomInset(top, right, bottom, left) {
  return { top, right, bottom, left };
}

function parseAdjFormula(formula) {
  if (formula == null) return null;
  const cleaned = formula.replace('val ', '');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

module.exports = {
  // ── SVG attribute helpers ─────────────────────────────────────────

  svgFillAttr(cssColor) {
    if (cssColor == null || cssColor === '' || cssColor === 'none') return 'fill="none"';
    const c = cssColor.replace(/^#/, '');
    if (c.length === 8) {
      const rgb = '#' + c.slice(0, 6);
      const alpha = parseInt(c.slice(6), 16);
      const opacity = alpha / 255.0;
      return `fill="${rgb}" fill-opacity="${f3(opacity)}"`;
    }
    return `fill="${cssColor}"`;
  },

  svgStopColorAttr(cssColor) {
    if (cssColor == null || cssColor === '') return 'stop-color="#000000"';
    const c = cssColor.replace(/^#/, '');
    if (c.length === 8) {
      const rgb = '#' + c.slice(0, 6);
      const alpha = parseInt(c.slice(6), 16);
      const opacity = alpha / 255.0;
      return `stop-color="${rgb}" stop-opacity="${f3(opacity)}"`;
    }
    return `stop-color="${cssColor}"`;
  },

  buildSvgStyleAttribute(styleValue) {
    return styleValue == null || styleValue.trim() === '' ? '' : ` style="${styleValue}"`;
  },

  combineSvgStyles(...styleValues) {
    let s = '';
    for (const styleValue of styleValues) {
      if (styleValue != null && styleValue.trim() !== '') s += styleValue;
    }
    return s;
  },

  buildSvgGroupDataAttributes(pptxShapeId) {
    // Deterministic, unique-per-slide handle: "<slideId>-<n>" (e.g. "1-1", "1_1-2").
    const elementId = `${this._currentSlideDataIdPrefix}-${++this._currentSlideDataIdCounter}`;
    if (pptxShapeId == null || pptxShapeId.trim() === '') return `data-id="${elementId}"`;
    return `data-id="${elementId}" data-pptx-id="${pptxShapeId}"`;
  },

  buildHiddenSvgStyle(pptxShapeId) {
    return pptxShapeId != null && pptxShapeId.trim() !== '' && this._currentHiddenShapeIds.has(pptxShapeId)
      ? 'display:none;'
      : '';
  },

  buildAnimatedSvgStyle(pptxShapeId) {
    if (pptxShapeId != null && pptxShapeId.trim() !== '' && this._currentAnimatedShapeStylesByShapeId.has(pptxShapeId)) {
      return this._currentAnimatedShapeStylesByShapeId.get(pptxShapeId);
    }
    return '';
  },

  getPptxShapeId(el) {
    for (const child of childElements(el)) {
      if (/^nv.*Pr$/.test(child.localName)) {
        const cNvPr = firstChild(child, 'cNvPr');
        if (cNvPr != null) {
          const id = attr(cNvPr, 'id');
          return id == null ? null : id;
        }
      }
    }
    return null;
  },

  // ── SVG Fill ──────────────────────────────────────────────────────

  getSvgFill(spPr, slidePart, width, height) {
    const solidFill = firstChild(spPr, 'solidFill');
    if (solidFill != null) {
      const color = this.getColorFromFill(solidFill);
      if (color !== '') return [this.svgFillAttr(color), ''];
    }

    const gradFill = firstChild(spPr, 'gradFill');
    if (gradFill != null) return this.getSvgGradientFill(gradFill);

    const blipFill = firstChild(spPr, 'blipFill');
    const imgPart = this.resolveImagePart(slidePart, firstChild(blipFill, 'blip'));
    if (imgPart != null) {
      const imgPath = this.saveImageFile(imgPart);
      return this.getSvgImageFill(imgPath, width, height);
    }

    const patternFill = firstChild(spPr, 'pattFill');
    if (patternFill != null) return this.getSvgPatternFill(patternFill);

    return ['fill="none"', ''];
  },

  getSvgGradientFill(gradFill) {
    const stops = elements(firstChild(gradFill, 'gsLst'), 'gs');
    if (stops.length < 2) return ['fill="none"', ''];

    const id = `grad${++this._defsCounter}`;
    const tryReadRectPctInt = (el, name) => {
      if (el == null) return null;
      const raw = attr(el, name);
      if (raw == null) return null;
      const v = parseInt(raw, 10);
      if (Number.isNaN(v)) return null;
      return v / 100000.0;
    };

    let ooxmlDeg = 0;
    const linFill = firstChild(gradFill, 'lin');
    let defs = '';

    const pathFill = firstChild(gradFill, 'path');
    let gradientTag = 'linearGradient';
    if (pathFill != null) {
      const fillRect = firstChild(pathFill, 'fillToRect');
      const left = tryReadRectPctInt(fillRect, 'l');
      const top = tryReadRectPctInt(fillRect, 't');
      const right = tryReadRectPctInt(fillRect, 'r');
      const bottom = tryReadRectPctInt(fillRect, 'b');

      const cx = left != null ? left : right != null ? 1.0 - right : 0.5;
      const cy = top != null ? top : bottom != null ? 1.0 - bottom : 0.5;

      let r = Math.max(
        Math.max(
          Math.sqrt((cx - 0) * (cx - 0) + (cy - 0) * (cy - 0)),
          Math.sqrt((cx - 1) * (cx - 1) + (cy - 0) * (cy - 0))
        ),
        Math.max(
          Math.sqrt((cx - 0) * (cx - 0) + (cy - 1) * (cy - 1)),
          Math.sqrt((cx - 1) * (cx - 1) + (cy - 1) * (cy - 1))
        )
      );
      r = Math.max(0.001, r);

      defs += `<radialGradient id="${id}" cx="${f4(cx)}" cy="${f4(cy)}" r="${f4(r)}">`;
      gradientTag = 'radialGradient';
    } else {
      if (linFill != null && attr(linFill, 'ang') != null) ooxmlDeg = numOr(linFill, 'ang') / 60000.0;

      const theta = (ooxmlDeg * Math.PI) / 180.0;
      const x1 = 0.5 - 0.5 * Math.cos(theta);
      const y1 = 0.5 - 0.5 * Math.sin(theta);
      const x2 = 0.5 + 0.5 * Math.cos(theta);
      const y2 = 0.5 + 0.5 * Math.sin(theta);
      defs += `<linearGradient id="${id}" x1="${f4(x1)}" y1="${f4(y1)}" x2="${f4(x2)}" y2="${f4(y2)}">`;
    }

    for (const stop of stops) {
      const color = this.getColorFromColorType(stop);
      const posPct = (attr(stop, 'pos') != null ? numOr(stop, 'pos') : 0) / 1000.0;
      defs += `<stop offset="${f1(posPct)}%" ${this.svgStopColorAttr(color)}/>`;
    }
    defs += `</${gradientTag}>`;

    return [`fill="url(#${id})"`, defs];
  },

  getSvgImageFill(imgPath, width, height) {
    const id = `imgPat${++this._defsCounter}`;
    const defs =
      `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${f1(width)}" height="${f1(height)}">` +
      `<image href="${imgPath}" width="${f1(width)}" height="${f1(height)}" preserveAspectRatio="none"/>` +
      '</pattern>';
    return [`fill="url(#${id})"`, defs];
  },

  getSvgPatternFill(patternFill) {
    const fg = '#' + this.resolveColorElement(firstChild(patternFill, 'fgClr'));
    const bg = '#' + this.resolveColorElement(firstChild(patternFill, 'bgClr'));
    const preset = attr(patternFill, 'prst') || '';

    const id = `pat${++this._defsCounter}`;
    const fullSvg = this.generatePatternSvg(preset, fg, bg);
    const start = fullSvg.indexOf('>') + 1;
    const end = fullSvg.lastIndexOf('</svg>');
    const content = start > 0 && end > start ? fullSvg.slice(start, end) : `<rect width='8' height='8' fill='${bg}'/>`;

    const defs = `<pattern id="${id}" patternUnits="userSpaceOnUse" width="8" height="8">` + content + '</pattern>';
    return [`fill="url(#${id})"`, defs];
  },

  // ── SVG Stroke ────────────────────────────────────────────────────

  getSvgStroke(spPr, width, height) {
    const outline = firstChild(spPr, 'ln');
    if (outline == null) return ['stroke="none"', ''];
    if (firstChild(outline, 'noFill') != null) return ['stroke="none"', ''];

    const solidFill = firstChild(outline, 'solidFill');
    if (solidFill == null) return ['stroke="none"', ''];

    const widthEmu = attr(outline, 'w') != null ? numOr(outline, 'w') : 12700;
    const strokePx = this.emuToPx(widthEmu);
    if (strokePx < 0.01) return ['stroke="none"', ''];

    let color = this.getColorFromFill(solidFill);
    if (color === '') color = '#000000';

    const cmpd = attr(outline, 'cmpd') || '';

    const isCompound = ciEq(cmpd, 'dbl') || ciEq(cmpd, 'thickThin') || ciEq(cmpd, 'thinThick') || ciEq(cmpd, 'tri');

    if (isCompound) {
      const overlay = this.buildCompoundBorderRects(cmpd, color, strokePx, width, height);
      return ['stroke="none"', overlay];
    }

    let dashAttr = '';
    const dash = firstChild(outline, 'prstDash');
    if (dash != null && attr(dash, 'val') != null) {
      const v = attr(dash, 'val') || '';
      const a = Math.max(1, strokePx);
      if (ciIncludes(v, 'dot')) dashAttr = ` stroke-dasharray="${f1(a)},${f1(2 * a)}"`;
      else if (!ciEq(v, 'solid')) dashAttr = ` stroke-dasharray="${f1(4 * a)},${f1(3 * a)}"`;
    }

    const strokeAttr = `stroke="${color}" stroke-width="${f2(strokePx)}"${dashAttr}`;
    return [strokeAttr, ''];
  },

  buildCompoundBorderRects(cmpd, color, strokePx, w, h) {
    const half = strokePx / 2.0;
    let svg = '';

    if (ciEq(cmpd, 'thickThin')) {
      const thick = strokePx * 0.5;
      const gap = strokePx * 0.25;
      const thin = strokePx * 0.25;
      const outerC = -half + thick / 2.0;
      const gapC = -half + thick + gap / 2.0;
      const innerC = -half + thick + gap + thin / 2.0;
      svg += `<rect x="${f1(gapC)}" y="${f1(gapC)}" width="${f1(w - 2 * gapC)}" height="${f1(h - 2 * gapC)}" fill="none" stroke="#FFFFFF" stroke-width="${f1(gap)}"/>`;
      svg += `<rect x="${f1(outerC)}" y="${f1(outerC)}" width="${f1(w - 2 * outerC)}" height="${f1(h - 2 * outerC)}" fill="none" stroke="${color}" stroke-width="${f1(thick)}"/>`;
      svg += `<rect x="${f1(innerC)}" y="${f1(innerC)}" width="${f1(w - 2 * innerC)}" height="${f1(h - 2 * innerC)}" fill="none" stroke="${color}" stroke-width="${f1(thin)}"/>`;
    } else if (ciEq(cmpd, 'thinThick')) {
      const thin = strokePx * 0.25;
      const gap = strokePx * 0.25;
      const thick = strokePx * 0.5;
      const outerC = -half + thin / 2.0;
      const gapC = -half + thin + gap / 2.0;
      const innerC = -half + thin + gap + thick / 2.0;
      svg += `<rect x="${f1(gapC)}" y="${f1(gapC)}" width="${f1(w - 2 * gapC)}" height="${f1(h - 2 * gapC)}" fill="none" stroke="#FFFFFF" stroke-width="${f1(gap)}"/>`;
      svg += `<rect x="${f1(outerC)}" y="${f1(outerC)}" width="${f1(w - 2 * outerC)}" height="${f1(h - 2 * outerC)}" fill="none" stroke="${color}" stroke-width="${f1(thin)}"/>`;
      svg += `<rect x="${f1(innerC)}" y="${f1(innerC)}" width="${f1(w - 2 * innerC)}" height="${f1(h - 2 * innerC)}" fill="none" stroke="${color}" stroke-width="${f1(thick)}"/>`;
    } else if (ciEq(cmpd, 'dbl')) {
      const line = strokePx / 3.0;
      const outerC = -half + line / 2.0;
      const gapC = 0.0;
      const innerC = half - line / 2.0;
      svg += `<rect x="${f1(gapC)}" y="${f1(gapC)}" width="${f1(w - 2 * gapC)}" height="${f1(h - 2 * gapC)}" fill="none" stroke="#FFFFFF" stroke-width="${f1(line)}"/>`;
      svg += `<rect x="${f1(outerC)}" y="${f1(outerC)}" width="${f1(w - 2 * outerC)}" height="${f1(h - 2 * outerC)}" fill="none" stroke="${color}" stroke-width="${f1(line)}"/>`;
      svg += `<rect x="${f1(innerC)}" y="${f1(innerC)}" width="${f1(w - 2 * innerC)}" height="${f1(h - 2 * innerC)}" fill="none" stroke="${color}" stroke-width="${f1(line)}"/>`;
    } else if (ciEq(cmpd, 'tri')) {
      const line = strokePx / 5.0;
      const outerC = -half + line / 2.0;
      const gap1C = -half + line + line / 2.0;
      const midC = 0.0;
      const gap2C = half - line - line / 2.0;
      const innerC = half - line / 2.0;
      svg += `<rect x="${f1(gap1C)}" y="${f1(gap1C)}" width="${f1(w - 2 * gap1C)}" height="${f1(h - 2 * gap1C)}" fill="none" stroke="#FFFFFF" stroke-width="${f1(line)}"/>`;
      svg += `<rect x="${f1(gap2C)}" y="${f1(gap2C)}" width="${f1(w - 2 * gap2C)}" height="${f1(h - 2 * gap2C)}" fill="none" stroke="#FFFFFF" stroke-width="${f1(line)}"/>`;
      svg += `<rect x="${f1(outerC)}" y="${f1(outerC)}" width="${f1(w - 2 * outerC)}" height="${f1(h - 2 * outerC)}" fill="none" stroke="${color}" stroke-width="${f1(line)}"/>`;
      svg += `<rect x="${f1(midC)}" y="${f1(midC)}" width="${f1(w - 2 * midC)}" height="${f1(h - 2 * midC)}" fill="none" stroke="${color}" stroke-width="${f1(line)}"/>`;
      svg += `<rect x="${f1(innerC)}" y="${f1(innerC)}" width="${f1(w - 2 * innerC)}" height="${f1(h - 2 * innerC)}" fill="none" stroke="${color}" stroke-width="${f1(line)}"/>`;
    }

    return svg;
  },

  // ── SVG Transform ────────────────────────────────────────────────

  getSvgTransform(xfrm, width, height, spPr = null) {
    const svgParts = [];

    if (xfrm != null) {
      const deg = numOr(xfrm, 'rot', 0) / 60000.0;
      const flipH = attr(xfrm, 'flipH') === '1' || ciEq(attr(xfrm, 'flipH'), 'true');
      const flipV = attr(xfrm, 'flipV') === '1' || ciEq(attr(xfrm, 'flipV'), 'true');

      let sx = 1;
      let sy = 1;
      if (flipV && !flipH) { sx = -1; sy = -1; }
      else { if (flipH) sx = -1; if (flipV) sy = -1; }

      const hasRot = Math.abs(deg) > 0.001;
      const hasFlip = sx !== 1 || sy !== 1;

      if (hasRot || hasFlip) {
        const cx = width / 2.0;
        const cy = height / 2.0;
        svgParts.push(`translate(${f1(cx)},${f1(cy)})`);
        if (hasRot) svgParts.push(`rotate(${f2(deg)})`);
        if (hasFlip) svgParts.push(`scale(${sx},${sy})`);
        svgParts.push(`translate(${f1(-cx)},${f1(-cy)})`);
      }
    }

    const svgTransform = svgParts.length > 0 ? ' ' + svgParts.join(' ') : '';

    let cssStyle = '';
    const scene3d = this.getScene3DTransform(spPr);
    if (scene3d !== '') cssStyle = `transform:${scene3d};transform-origin:center center;transform-style:preserve-3d;`;

    return [svgTransform, cssStyle];
  },

  // ── SVG Effects ──────────────────────────────────────────────────

  getSvgEffectsStyle(spPr) {
    const filters = [];

    const effects = firstChild(spPr, 'effectLst');
    if (effects != null) {
      const outer = firstChild(effects, 'outerShdw');
      if (outer != null) {
        const blur = this.emuToPx(numOr(outer, 'blurRad', 0));
        const dist = this.emuToPx(numOr(outer, 'dist', 0));
        const angle = ((numOr(outer, 'dir', 0) / 60000.0) * Math.PI) / 180.0;
        const dx = dist * Math.cos(angle);
        const dy = dist * Math.sin(angle);
        const color = this.getColorFromEffectColorContainer(outer, '00000080');
        filters.push(`drop-shadow(${f1(dx)}px ${f1(dy)}px ${f1(blur)}px ${color})`);
      }

      const glow = firstChild(effects, 'glow');
      if (glow != null) {
        const radius = this.emuToPx(numOr(glow, 'rad', 0));
        const color = this.getColorFromEffectColorContainer(glow, 'FFFFFF66');
        filters.push(`drop-shadow(0 0 ${f1(radius)}px ${color})`);
      }
    }

    if (filters.length === 0) return '';
    return `filter:${filters.join(' ')};`;
  },

  // ── Shape Rendering ──────────────────────────────────────────────

  renderShape(sb, defsSb, shape, slidePart) {
    const spPr = shapePropsOf(shape);
    const txBody = textBodyOf(shape);
    if (spPr == null) return;

    const ph = placeholderOf(shape);

    let [left, top, width, height] = this.getTransform(spPr);

    if (width <= 0 || height <= 0) {
      if (ph != null) {
        const resolved = this.resolvePhTransform(ph, slidePart);
        if (resolved != null) [left, top, width, height] = resolved;
      }
    }

    if (width <= 0 || height <= 0) return;

    let [fillAttr, fillDefs] = this.getSvgFill(spPr, slidePart, width, height);
    if (fillDefs !== '') defsSb.appendLine(fillDefs);

    let [strokeAttr, strokeOverlay] = this.getSvgStroke(spPr, width, height);

    const shapeStyle = firstChild(shape, 'style');
    if (shapeStyle != null) {
      if (
        fillAttr === 'fill="none"' &&
        firstChild(spPr, 'solidFill') == null &&
        firstChild(spPr, 'gradFill') == null &&
        firstChild(spPr, 'blipFill') == null &&
        firstChild(spPr, 'pattFill') == null &&
        firstChild(spPr, 'noFill') == null
      ) {
        fillAttr = this.resolveStyleFillRef(firstChild(shapeStyle, 'fillRef'));
      }
      if (strokeAttr === 'stroke="none"' && firstChild(spPr, 'ln') == null) {
        strokeAttr = this.resolveStyleLineRef(firstChild(shapeStyle, 'lnRef'));
      }
    }

    const presetGeom = this.getPresetGeometryData(spPr, width, height);
    const inset = presetGeom ? presetGeom.inset : geomInset(0, 0, 0, 0);

    let customFillElement = '';
    let customStrokeElement = '';
    let customClipId = '';
    if (presetGeom == null) {
      const custom = this.getCustomGeometrySvgElements(spPr, width, height, fillAttr, defsSb);
      customFillElement = custom.fillElement;
      customStrokeElement = custom.strokeElement;
      customClipId = custom.clipId;
    }

    const inherited = this.resolveInheritedProps(ph, slidePart);

    let vAlignStyle = '';
    const bodyPr = firstChild(txBody, 'bodyPr');
    let effectiveAnchor = inherited.verticalAnchor;
    if (bodyPr != null && attr(bodyPr, 'anchor') != null) {
      const av = attr(bodyPr, 'anchor');
      if (av === 'b') effectiveAnchor = 'b';
      else if (av === 'ctr') effectiveAnchor = 'ctr';
      else effectiveAnchor = 't';
    }
    if (effectiveAnchor === 'b') vAlignStyle = 'display:flex;flex-direction:column;justify-content:flex-end;';
    else if (effectiveAnchor === 'ctr') vAlignStyle = 'display:flex;flex-direction:column;justify-content:center;';

    let fontScale = 1.0;
    let lnSpcReduction = 0.0;
    const normAutofit = firstChild(bodyPr, 'normAutofit');
    if (normAutofit != null) {
      if (attr(normAutofit, 'fontScale') != null) fontScale = numOr(normAutofit, 'fontScale') / 100000.0;
      if (attr(normAutofit, 'lnSpcReduction') != null) lnSpcReduction = numOr(normAutofit, 'lnSpcReduction') / 100000.0;
    }

    const tInsEmu = attr(bodyPr, 'tIns') != null ? numOr(bodyPr, 'tIns') : 45720;
    const bInsEmu = attr(bodyPr, 'bIns') != null ? numOr(bodyPr, 'bIns') : 45720;
    const lInsEmu = attr(bodyPr, 'lIns') != null ? numOr(bodyPr, 'lIns') : 91440;
    const rInsEmu = attr(bodyPr, 'rIns') != null ? numOr(bodyPr, 'rIns') : 91440;
    const tInsPx = Math.max(this.emuToPx(tInsEmu), inset.top);
    const bInsPx = Math.max(this.emuToPx(bInsEmu), inset.bottom);
    const lInsPx = Math.max(this.emuToPx(lInsEmu), inset.left);
    const rInsPx = Math.max(this.emuToPx(rInsEmu), inset.right);
    const padStyle = `padding:${f1(tInsPx)}px ${f1(rInsPx)}px ${f1(bInsPx)}px ${f1(lInsPx)}px;`;

    const compensateLeading = effectiveAnchor !== 'ctr' && effectiveAnchor !== 'b';

    let vertTextStyle = '';
    if (bodyPr != null && attr(bodyPr, 'vert') != null) {
      const v = attr(bodyPr, 'vert');
      if (v === 'vert' || v === 'eaVert') vertTextStyle = 'writing-mode:vertical-rl;';
      else if (v === 'vert270') vertTextStyle = 'writing-mode:vertical-rl;transform:rotate(180deg);';
    }
    const isNoWrapText = bodyPr != null && attr(bodyPr, 'wrap') === 'none';
    const vertVal = bodyPr != null ? attr(bodyPr, 'vert') : null;
    const isVerticalText = vertVal === 'vert' || vertVal === 'eaVert' || vertVal === 'vert270';

    const [svgXform, css3dStyle] = this.getSvgTransform(firstChild(spPr, 'xfrm'), width, height, spPr);

    const effectsStyle = this.getSvgEffectsStyle(spPr);

    const pptxShapeId = this.getPptxShapeId(shape) || '';
    const groupDataAttributes = this.buildSvgGroupDataAttributes(pptxShapeId);
    const hiddenStyle = this.buildHiddenSvgStyle(pptxShapeId);
    const innerStyle = this.combineSvgStyles(css3dStyle, effectsStyle);
    const animationStyle = this.buildAnimatedSvgStyle(pptxShapeId);
    const morphOverlay = this._currentMorphPictureOverlaysByShapeId.get(pptxShapeId);
    const wrapperStyle = this.buildSvgStyleAttribute(hiddenStyle);
    const animatedWrapperStyle = this.buildSvgStyleAttribute(this.combineSvgStyles(hiddenStyle, animationStyle));
    const directStyle = this.buildSvgStyleAttribute(this.combineSvgStyles(hiddenStyle, innerStyle));
    const innerStyleAttr = this.buildSvgStyleAttribute(innerStyle);

    const defaultTextColor = `color:${this.getThemeTextColor()};`;

    const needsStaticWrapper = innerStyle !== '';
    const needsAnimationWrapper = animationStyle !== '';
    if (needsAnimationWrapper) {
      sb.appendLine(`<g ${groupDataAttributes}${animatedWrapperStyle}>`);
      sb.appendLine(`<g transform="translate(${f0(left)},${f0(top)})${svgXform}">`);
      if (needsStaticWrapper) sb.appendLine(`<g${innerStyleAttr}>`);
    } else if (needsStaticWrapper) {
      sb.appendLine(`<g ${groupDataAttributes} transform="translate(${f0(left)},${f0(top)})${svgXform}"${wrapperStyle}>`);
      sb.appendLine(`<g${innerStyleAttr}>`);
    } else {
      sb.appendLine(`<g ${groupDataAttributes} transform="translate(${f0(left)},${f0(top)})${svgXform}"${directStyle}>`);
    }

    if (presetGeom != null) {
      if (presetGeom.type === 'polygon') sb.appendLine(`<polygon points="${presetGeom.data}" ${fillAttr} ${strokeAttr}/>`);
      else if (presetGeom.type === 'roundrect')
        sb.appendLine(`<rect width="${f1(width)}" height="${f1(height)}" rx="${f1(presetGeom.rx)}" ry="${f1(presetGeom.ry)}" ${fillAttr} ${strokeAttr}/>`);
      else if (presetGeom.type === 'ellipse')
        sb.appendLine(`<ellipse cx="${f1(width / 2.0)}" cy="${f1(height / 2.0)}" rx="${f1(width / 2.0)}" ry="${f1(height / 2.0)}" ${fillAttr} ${strokeAttr}/>`);
      else if (presetGeom.type === 'path') sb.appendLine(`<path d="${presetGeom.data}" ${fillAttr} ${strokeAttr}/>`);
    } else if (customFillElement !== '') {
      sb.appendLine(customFillElement);
      if (customStrokeElement !== '') sb.appendLine(customStrokeElement);
    } else {
      sb.appendLine(`<rect width="${f1(width)}" height="${f1(height)}" ${fillAttr} ${strokeAttr}/>`);
    }

    if (strokeOverlay !== '') sb.appendLine(strokeOverlay);

    if (txBody != null) {
      const clipAttr = customClipId !== '' ? ` clip-path="url(#${customClipId})"` : '';
      const wrapStyle = isNoWrapText
        ? 'white-space:nowrap;word-wrap:normal;overflow-wrap:normal;'
        : 'word-wrap:break-word;overflow-wrap:break-word;';
      sb.appendLine(`<foreignObject x="0" y="0" width="${f1(width)}" height="${f1(height)}" overflow="visible"${clipAttr}>`);
      sb.appendLine(
        `<div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;${wrapStyle}overflow:visible;${defaultTextColor}${vAlignStyle}${padStyle}${vertTextStyle}">`
      );
      this.renderTextBody(sb, txBody, inherited, fontScale, lnSpcReduction, compensateLeading, isNoWrapText, isVerticalText);
      sb.appendLine('</div>');
      sb.appendLine('</foreignObject>');
    }

    if (morphOverlay != null && morphOverlay.previousSvgMarkup && morphOverlay.previousSvgMarkup.trim() !== '') {
      sb.append(morphOverlay.previousSvgMarkup);
    }

    if (needsStaticWrapper) sb.appendLine('</g>');
    sb.appendLine('</g>');
    if (needsAnimationWrapper) sb.appendLine('</g>');
  },

  renderPicture(sb, defsSb, picture, slidePart) {
    const spPr = shapePropsOf(picture);
    if (spPr == null) return;

    let [left, top, width, height] = this.getTransform(spPr);
    if (width <= 0 || height <= 0) return;

    const nvPr = firstChild(firstChild(picture, 'nvPicPr'), 'nvPr');
    const videoFile = firstChild(nvPr, 'videoFile');
    const audioFile = firstChild(nvPr, 'audioFile');

    let mediaSrc = null;
    let mediaType = null;
    const isRelNs = (ns) => ns === NS.R || (ns || '').toLowerCase().includes('relationships');
    if (videoFile != null && attr(videoFile, 'link', isRelNs) != null) {
      mediaSrc = this.resolveMediaPath(slidePart, attr(videoFile, 'link', isRelNs));
      mediaType = 'video';
    } else if (audioFile != null && attr(audioFile, 'link', isRelNs) != null) {
      mediaSrc = this.resolveMediaPath(slidePart, attr(audioFile, 'link', isRelNs));
      mediaType = 'audio';
    }

    if (mediaSrc == null && (videoFile != null || audioFile != null)) {
      mediaSrc = this.resolveP14Media(nvPr, slidePart);
      if (mediaSrc != null) mediaType = videoFile != null ? 'video' : 'audio';
    }

    let imgSrc = null;
    const blipFill = firstChild(picture, 'blipFill');
    const blip = firstChild(blipFill, 'blip');
    const imgPart = this.resolveImagePart(slidePart, blip);
    if (imgPart != null) imgSrc = this.saveImageFile(imgPart);

    const [fillAttr, fillDefs] = this.getSvgFill(spPr, slidePart, width, height);
    if (fillDefs !== '') defsSb.appendLine(fillDefs);

    const [strokeAttr, strokeOverlay] = this.getSvgStroke(spPr, width, height);

    const presetGeom = this.getPresetGeometryData(spPr, width, height);
    let customFillElement = '';
    let customStrokeElement = '';
    let customClipId = '';
    if (presetGeom == null) {
      const custom = this.getCustomGeometrySvgElements(spPr, width, height, fillAttr, defsSb);
      customFillElement = custom.fillElement;
      customStrokeElement = custom.strokeElement;
      customClipId = custom.clipId;
    }

    let clipAttr = '';
    if (presetGeom != null && presetGeom.type !== 'rect') {
      const clipId = this.appendPresetGeometryClip(defsSb, presetGeom, width, height);
      if (clipId !== '') clipAttr = ` clip-path="url(#${clipId})"`;
    } else if (customClipId !== '') {
      clipAttr = ` clip-path="url(#${customClipId})"`;
    }

    const [svgXform, css3dStyle] = this.getSvgTransform(firstChild(spPr, 'xfrm'), width, height, spPr);
    const effectsStyle = this.getSvgEffectsStyle(spPr);
    const pptxShapeId = this.getPptxShapeId(picture) || '';
    const groupDataAttributes = this.buildSvgGroupDataAttributes(pptxShapeId);
    const hiddenStyle = this.buildHiddenSvgStyle(pptxShapeId);
    const innerStyle = this.combineSvgStyles(css3dStyle, effectsStyle);
    const animationStyle = this.buildAnimatedSvgStyle(pptxShapeId);
    const morphPictureOverlay = this._currentMorphPictureOverlaysByShapeId.get(pptxShapeId);
    const wrapperStyle = this.buildSvgStyleAttribute(hiddenStyle);
    const animatedWrapperStyle = this.buildSvgStyleAttribute(this.combineSvgStyles(hiddenStyle, animationStyle));
    const directStyle = this.buildSvgStyleAttribute(this.combineSvgStyles(hiddenStyle, innerStyle));
    const innerStyleAttr = this.buildSvgStyleAttribute(innerStyle);

    const needsStaticWrapper = innerStyle !== '';
    const needsAnimationWrapper = animationStyle !== '';
    if (needsAnimationWrapper) {
      sb.appendLine(`<g ${groupDataAttributes}${animatedWrapperStyle}>`);
      sb.appendLine(`<g transform="translate(${f0(left)},${f0(top)})${svgXform}">`);
      if (needsStaticWrapper) sb.appendLine(`<g${innerStyleAttr}>`);
    } else if (needsStaticWrapper) {
      sb.appendLine(`<g ${groupDataAttributes} transform="translate(${f0(left)},${f0(top)})${svgXform}"${wrapperStyle}>`);
      sb.appendLine(`<g${innerStyleAttr}>`);
    } else {
      sb.appendLine(`<g ${groupDataAttributes} transform="translate(${f0(left)},${f0(top)})${svgXform}"${directStyle}>`);
    }

    if (fillAttr !== 'fill="none"') {
      if (presetGeom != null) this.appendPresetGeometryElement(sb, presetGeom, width, height, fillAttr, 'stroke="none"');
      else if (customFillElement !== '') sb.appendLine(customFillElement);
      else sb.appendLine(`<rect width="${f1(width)}" height="${f1(height)}" ${fillAttr} stroke="none"/>`);
    }

    if (mediaSrc != null && mediaType === 'video') {
      const posterAttr = imgSrc != null ? ` poster="${imgSrc}"` : '';
      sb.appendLine(`<foreignObject x="0" y="0" width="${f1(width)}" height="${f1(height)}"${clipAttr}>`);
      sb.appendLine(`<video xmlns="http://www.w3.org/1999/xhtml" src="${mediaSrc}" style="width:100%;height:100%;object-fit:fill;" controls${posterAttr}></video>`);
      sb.appendLine('</foreignObject>');
    } else if (mediaSrc != null && mediaType === 'audio') {
      sb.appendLine(`<foreignObject x="0" y="0" width="${f1(width)}" height="${f1(height)}"${clipAttr}>`);
      sb.appendLine('<div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;width:100%;height:100%;">');
      if (imgSrc != null) sb.appendLine(`<img src="${imgSrc}" style="width:100%;height:100%;object-fit:fill;" alt="audio" />`);
      sb.appendLine(`<audio src="${mediaSrc}" style="width:100%;position:absolute;bottom:0;left:0;" controls></audio>`);
      sb.appendLine('</div>');
      sb.appendLine('</foreignObject>');
    } else if (imgSrc != null) {
      sb.appendLine(`<image href="${imgSrc}" width="${f1(width)}" height="${f1(height)}" preserveAspectRatio="none"${clipAttr}/>`);
      if (morphPictureOverlay != null && morphPictureOverlay.previousImagePath && morphPictureOverlay.previousImagePath.trim() !== '') {
        this._currentSlideImages.push(morphPictureOverlay.previousImagePath);
        const overlayStyleAttr = this.buildSvgStyleAttribute(morphPictureOverlay.previousImageStyle);
        sb.appendLine(
          `<image href="${morphPictureOverlay.previousImagePath}" width="${f1(width)}" height="${f1(height)}" preserveAspectRatio="none"${clipAttr}${overlayStyleAttr}/>`
        );
      }
    }

    if (strokeAttr !== 'stroke="none"') {
      if (presetGeom != null) this.appendPresetGeometryElement(sb, presetGeom, width, height, 'fill="none"', strokeAttr);
      else if (customStrokeElement !== '') sb.appendLine(customStrokeElement);
      else sb.appendLine(`<rect width="${f1(width)}" height="${f1(height)}" fill="none" ${strokeAttr}/>`);
    }

    if (strokeOverlay !== '') sb.appendLine(strokeOverlay);

    if (needsStaticWrapper) sb.appendLine('</g>');
    sb.appendLine('</g>');
    if (needsAnimationWrapper) sb.appendLine('</g>');
  },

  appendPresetGeometryElement(sb, presetGeom, width, height, fillAttr, strokeAttr) {
    if (presetGeom.type === 'polygon') sb.appendLine(`<polygon points="${presetGeom.data}" ${fillAttr} ${strokeAttr}/>`);
    else if (presetGeom.type === 'roundrect')
      sb.appendLine(`<rect width="${f1(width)}" height="${f1(height)}" rx="${f1(presetGeom.rx)}" ry="${f1(presetGeom.ry)}" ${fillAttr} ${strokeAttr}/>`);
    else if (presetGeom.type === 'ellipse')
      sb.appendLine(`<ellipse cx="${f1(width / 2.0)}" cy="${f1(height / 2.0)}" rx="${f1(width / 2.0)}" ry="${f1(height / 2.0)}" ${fillAttr} ${strokeAttr}/>`);
    else if (presetGeom.type === 'path') sb.appendLine(`<path d="${presetGeom.data}" ${fillAttr} ${strokeAttr}/>`);
    else sb.appendLine(`<rect width="${f1(width)}" height="${f1(height)}" ${fillAttr} ${strokeAttr}/>`);
  },

  appendPresetGeometryClip(defsSb, presetGeom, width, height) {
    const clipId = `clip${++this._defsCounter}`;
    defsSb.appendLine(`<clipPath id="${clipId}">`);
    if (presetGeom.type === 'polygon') defsSb.appendLine(`<polygon points="${presetGeom.data}"/>`);
    else if (presetGeom.type === 'roundrect')
      defsSb.appendLine(`<rect width="${f1(width)}" height="${f1(height)}" rx="${f1(presetGeom.rx)}" ry="${f1(presetGeom.ry)}"/>`);
    else if (presetGeom.type === 'ellipse')
      defsSb.appendLine(`<ellipse cx="${f1(width / 2.0)}" cy="${f1(height / 2.0)}" rx="${f1(width / 2.0)}" ry="${f1(height / 2.0)}"/>`);
    else if (presetGeom.type === 'path') defsSb.appendLine(`<path d="${presetGeom.data}"/>`);
    else {
      defsSb.appendLine('</clipPath>');
      return '';
    }
    defsSb.appendLine('</clipPath>');
    return clipId;
  },

  resolveMediaPath(slidePart, relId) {
    for (const dpRef of slidePart.dataPartReferenceRelationships()) {
      if (dpRef.id === relId) return this.saveMediaFile(dpRef.dataPart);
    }
    return null;
  },

  resolveP14Media(nvPr, slidePart) {
    const extList = firstChild(nvPr, 'extLst');
    if (extList == null) return null;

    for (const ext of elements(extList, 'ext')) {
      for (const child of childElements(ext)) {
        const embedAttr = attr(child, 'embed', (ns) => ns === NS.R || (ns || '').includes('relationships'));
        if (embedAttr != null) {
          for (const dpRef of slidePart.dataPartReferenceRelationships()) {
            if (dpRef.id === embedAttr) return this.saveMediaFile(dpRef.dataPart);
          }
        }
      }
    }
    return null;
  },

  renderGraphicFrame(sb, defsSb, frame, slidePart) {
    const xfrm = firstChild(frame, 'xfrm');
    if (xfrm == null) return;

    const off = firstChild(xfrm, 'off');
    const ext = firstChild(xfrm, 'ext');
    const left = this.emuToPx(attrLong(off, 'x'));
    const top = this.emuToPx(attrLong(off, 'y'));
    const width = this.emuToPx(attrLong(ext, 'cx'));
    const height = this.emuToPx(attrLong(ext, 'cy'));
    const pptxShapeId = this.getPptxShapeId(frame) || '';
    const groupDataAttributes = this.buildSvgGroupDataAttributes(pptxShapeId);
    const hiddenStyleAttr = this.buildSvgStyleAttribute(this.buildHiddenSvgStyle(pptxShapeId));

    const graphic = descendants(frame, 'graphic', A)[0];
    const graphicData = firstChild(graphic, 'graphicData');
    const table = graphicData ? descendants(graphicData, 'tbl', A)[0] : null;

    if (table != null) {
      sb.appendLine(`<g ${groupDataAttributes} transform="translate(${f0(left)},${f0(top)})"${hiddenStyleAttr}>`);
      sb.appendLine(`<foreignObject x="0" y="0" width="${f1(width)}" height="${f1(height)}">`);
      sb.appendLine('<div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;">');
      this.renderTable(sb, table);
      sb.appendLine('</div>');
      sb.appendLine('</foreignObject>');
      sb.appendLine('</g>');
      return;
    }

    const chartRefEl = graphicData
      ? descendants(graphicData).find((e) => e.localName === 'chart' && e.namespaceURI === C)
      : null;
    const chartRelId = chartRefEl
      ? attr(chartRefEl, 'id', (ns) => (ns || '').toLowerCase().includes('relationships')) || ''
      : '';
    if (chartRelId !== '') {
      const chartPart = slidePart.getPartById(chartRelId);
      if (chartPart != null && chartPart.contentType && chartPart.contentType.includes('chart')) {
        this.renderChartFrame(sb, chartPart, left, top, width, height, pptxShapeId);
      }
    }
  },

  renderChartFrame(sb, chartPart, left, top, width, height, pptxShapeId) {
    const esc = (s) =>
      (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    const root = chartPart.root;

    const barChart = descendants(root, 'barChart', C)[0];
    if (barChart == null) return;

    const serEls = elements(barChart, 'ser', C);
    if (serEls.length === 0) return;

    const readStringPoints = (container) => {
      if (container == null) return [];
      return descendants(container, 'pt', C)
        .sort((p1, p2) => (numOr(p1, 'idx', 0) || 0) - (numOr(p2, 'idx', 0) || 0))
        .map((p) => innerText(firstChild(p, 'v')) || '');
    };

    const readNumericPoints = (container) => {
      const result = [];
      if (container == null) return result;
      for (const pt of descendants(container, 'pt', C).sort((p1, p2) => (numOr(p1, 'idx', 0) || 0) - (numOr(p2, 'idx', 0) || 0))) {
        const v = innerText(firstChild(pt, 'v')) || '0';
        const n = Number(v);
        result.push(Number.isFinite(n) ? n : 0);
      }
      return result;
    };

    const resolveChartSchemeColor = (token) => {
      let key = token;
      if (this._colorMap.has(key)) key = this._colorMap.get(key);
      if (this._themeColors.has(key)) return '#' + this._themeColors.get(key);
      if (key === 'accent1') return '#5B9BD5';
      if (key === 'accent2') return '#ED7D31';
      if (key === 'accent3') return '#A5A5A5';
      if (key === 'accent4') return '#FFC000';
      if (key === 'accent5') return '#4472C4';
      if (key === 'accent6') return '#70AD47';
      if (key === 'tx1' || key === 'dk1') return '#222222';
      if (key === 'lt1' || key === 'bg1') return '#FFFFFF';
      return '#5B9BD5';
    };

    const resolveSeriesColor = (ser, index) => {
      const solidFill = firstChild(firstChild(ser, 'spPr'), 'solidFill');
      const srgb = attr(firstChild(solidFill, 'srgbClr'), 'val');
      if (srgb && srgb.trim() !== '') return '#' + srgb;
      const scheme = attr(firstChild(solidFill, 'schemeClr'), 'val');
      if (scheme && scheme.trim() !== '') return resolveChartSchemeColor(scheme);
      const fallback = ['#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000', '#4472C4', '#70AD47'];
      return fallback[index % fallback.length];
    };

    const readSeriesName = (ser) => {
      const txPt = descendants(firstChild(ser, 'tx'), 'pt', C)[0];
      return (innerText(firstChild(txPt, 'v')) || '') || `Series ${serEls.indexOf(ser) + 1}`;
    };

    let categories = readStringPoints(firstChild(serEls[0], 'cat'));
    if (categories.length === 0) categories = [1, 2, 3, 4].map((i) => `Category ${i}`);

    const seriesList = [];
    for (let i = 0; i < serEls.length; i++) {
      const ser = serEls[i];
      let vals = readNumericPoints(firstChild(ser, 'val'));
      if (vals.length === 0) continue;
      while (vals.length < categories.length) vals.push(0);
      if (vals.length > categories.length) vals = vals.slice(0, categories.length);
      seriesList.push({ name: readSeriesName(ser), values: vals, color: resolveSeriesColor(ser, i) });
    }
    if (seriesList.length === 0) return;

    const allValues = seriesList.flatMap((s) => s.values);
    const maxVal = Math.max(1.0, allValues.length ? Math.max(...allValues) : 1);
    let axisMax = Math.ceil(maxVal);
    if (axisMax - maxVal < 0.25) axisMax += 1;
    const tickCount = Math.trunc(Math.min(8, Math.max(4, axisMax)));
    const tickStep = axisMax / tickCount;

    const chartPadLeft = Math.max(42, width * 0.06);
    const chartPadRight = Math.max(22, width * 0.04);
    const chartPadTop = Math.max(12, height * 0.04);
    const legendH = Math.max(28, height * 0.1);
    const chartPadBottom = legendH + Math.max(28, height * 0.08);

    const plotX = chartPadLeft;
    const plotY = chartPadTop;
    const plotW = Math.max(60, width - chartPadLeft - chartPadRight);
    const plotH = Math.max(40, height - chartPadTop - chartPadBottom);

    const gridColor = '#FFFFFF';
    const axisLabelColor = '#FFFFFF';
    const legendColor = '#FFFFFF';
    const groupDataAttributes = this.buildSvgGroupDataAttributes(pptxShapeId);
    const hiddenStyleAttr = this.buildSvgStyleAttribute(this.buildHiddenSvgStyle(pptxShapeId));

    sb.appendLine(`<g ${groupDataAttributes} transform="translate(${f0(left)},${f0(top)})"${hiddenStyleAttr}>`);

    for (let i = 0; i <= tickCount; i++) {
      const value = i * tickStep;
      const y = plotY + plotH - (value / axisMax) * plotH;
      sb.appendLine(`<line x1="${f1(plotX)}" y1="${f1(y)}" x2="${f1(plotX + plotW)}" y2="${f1(y)}" stroke="${gridColor}" stroke-width="1"/>`);
      sb.appendLine(`<text x="${f1(plotX - 12)}" y="${f1(y + 4)}" text-anchor="end" font-size="12" fill="${axisLabelColor}" font-family="Tahoma">${f0(value)}</text>`);
    }

    const catCount = categories.length;
    const serCount = seriesList.length;
    const slotW = plotW / Math.max(1, catCount);
    const clusterW = slotW * 0.62;
    const groupGap = Math.max(1.0, (clusterW * 0.08) / Math.max(1, serCount));
    const barW = Math.max(2.0, (clusterW - (serCount - 1) * groupGap) / Math.max(1, serCount));
    for (let cIdx = 0; cIdx < catCount; cIdx++) {
      const clusterX = plotX + cIdx * slotW + (slotW - clusterW) / 2.0;
      for (let sIdx = 0; sIdx < serCount; sIdx++) {
        const v = seriesList[sIdx].values[cIdx];
        const barH = Math.max(0, (v / axisMax) * plotH);
        const x = clusterX + sIdx * (barW + groupGap);
        const y = plotY + plotH - barH;
        sb.appendLine(`<rect x="${f1(x)}" y="${f1(y)}" width="${f1(barW)}" height="${f1(barH)}" fill="${seriesList[sIdx].color}"/>`);
      }
    }

    for (let cIdx = 0; cIdx < catCount; cIdx++) {
      const x = plotX + cIdx * slotW + slotW / 2.0;
      const y = plotY + plotH + 26;
      sb.appendLine(`<text x="${f1(x)}" y="${f1(y)}" text-anchor="middle" font-size="12" fill="${axisLabelColor}" font-family="Tahoma">${esc(categories[cIdx])}</text>`);
    }

    const legendY = height - Math.max(8, legendH * 0.3);
    const itemGap = 14;
    const swatchW = 10;
    const textWApprox = 52;
    const itemW = swatchW + 6 + textWApprox;
    const legendW = serCount * itemW + (serCount - 1) * itemGap;
    const legendX = Math.max(0, (width - legendW) / 2.0);
    for (let i = 0; i < serCount; i++) {
      const x = legendX + i * (itemW + itemGap);
      sb.appendLine(`<rect x="${f1(x)}" y="${f1(legendY - 9)}" width="${f1(swatchW)}" height="${f1(swatchW)}" fill="${seriesList[i].color}"/>`);
      sb.appendLine(`<text x="${f1(x + swatchW + 6)}" y="${f1(legendY)}" font-size="12" fill="${legendColor}" font-family="Tahoma">${esc(seriesList[i].name)}</text>`);
    }

    sb.appendLine('</g>');
  },

  renderConnectionShape(sb, defsSb, cxnShape, slidePart) {
    const spPr = shapePropsOf(cxnShape);
    if (spPr == null) return;

    const xfrm = firstChild(spPr, 'xfrm');
    if (xfrm == null) return;

    const off = firstChild(xfrm, 'off');
    const ext = firstChild(xfrm, 'ext');
    const left = this.emuToPx(attrLong(off, 'x'));
    const top = this.emuToPx(attrLong(off, 'y'));
    const width = this.emuToPx(attrLong(ext, 'cx'));
    const height = this.emuToPx(attrLong(ext, 'cy'));
    const pptxShapeId = this.getPptxShapeId(cxnShape) || '';
    const groupDataAttributes = this.buildSvgGroupDataAttributes(pptxShapeId);
    const hiddenStyleAttr = this.buildSvgStyleAttribute(this.buildHiddenSvgStyle(pptxShapeId));

    const flipH = attr(xfrm, 'flipH') === '1' || ciEq(attr(xfrm, 'flipH'), 'true');
    const flipV = attr(xfrm, 'flipV') === '1' || ciEq(attr(xfrm, 'flipV'), 'true');

    const outline = firstChild(spPr, 'ln');
    if (outline == null) {
      const shapeStyle = firstChild(cxnShape, 'style');
      const lnRef = firstChild(shapeStyle, 'lnRef');
      if (lnRef != null) {
        const lnRefAttr = this.resolveStyleLineRef(lnRef);
        if (lnRefAttr !== 'stroke="none"') {
          sb.appendLine(`<g ${groupDataAttributes} transform="translate(${f0(left)},${f0(top)})"${hiddenStyleAttr}>`);
          sb.appendLine(`<line x1="0" y1="0" x2="${f1(width)}" y2="${f1(height)}" ${lnRefAttr}/>`);
          sb.appendLine('</g>');
        }
      }
      return;
    }
    if (firstChild(outline, 'noFill') != null) return;

    const widthEmu = attr(outline, 'w') != null ? numOr(outline, 'w') : 12700;
    const strokePx = this.emuToPx(widthEmu);
    if (strokePx < 0.01) return;

    let strokeColor = '#000000';
    const solidFill = firstChild(outline, 'solidFill');
    if (solidFill != null) {
      const c = this.getColorFromFill(solidFill);
      if (c !== '') strokeColor = c;
    }

    let dashAttr = '';
    const dash = firstChild(outline, 'prstDash');
    if (dash != null && attr(dash, 'val') != null) {
      const v = attr(dash, 'val') || '';
      const a = Math.max(1, strokePx);
      if (ciIncludes(v, 'dot')) dashAttr = ` stroke-dasharray="${f1(a)},${f1(2 * a)}"`;
      else if (!ciEq(v, 'solid')) dashAttr = ` stroke-dasharray="${f1(4 * a)},${f1(3 * a)}"`;
    }

    const headEnd = firstChild(outline, 'headEnd');
    const tailEnd = firstChild(outline, 'tailEnd');

    let headMarkerId = '';
    let tailMarkerId = '';

    if (headEnd != null && attr(headEnd, 'type') != null && attr(headEnd, 'type') !== 'none') {
      headMarkerId = `marker${++this._defsCounter}`;
      defsSb.appendLine(this.buildArrowMarkerDef(headMarkerId, attr(headEnd, 'type'), strokeColor, strokePx, true));
    }
    if (tailEnd != null && attr(tailEnd, 'type') != null && attr(tailEnd, 'type') !== 'none') {
      tailMarkerId = `marker${++this._defsCounter}`;
      defsSb.appendLine(this.buildArrowMarkerDef(tailMarkerId, attr(tailEnd, 'type'), strokeColor, strokePx, false));
    }

    let markerAttrs = '';
    if (headMarkerId !== '') markerAttrs += ` marker-start="url(#${headMarkerId})"`;
    if (tailMarkerId !== '') markerAttrs += ` marker-end="url(#${tailMarkerId})"`;

    const prstGeom = firstChild(spPr, 'prstGeom');
    let presetName = prstGeom ? attr(prstGeom, 'prst') || '' : '';

    sb.appendLine(`<g ${groupDataAttributes} transform="translate(${f0(left)},${f0(top)})"${hiddenStyleAttr}>`);

    const x1 = flipH ? width : 0;
    const y1 = flipV ? height : 0;
    const x2 = flipH ? 0 : width;
    const y2 = flipV ? 0 : height;

    if (ciIncludes(presetName, 'bentConnector')) {
      const avLst = firstChild(prstGeom, 'avLst');
      let adjVal = 50000;
      const guide = firstChild(avLst, 'gd');
      if (guide != null) {
        const parsed = parseFloat((attr(guide, 'fmla') || '').replace('val ', ''));
        if (!Number.isNaN(parsed)) adjVal = parsed;
      }

      const midFraction = adjVal / 100000.0;
      let midX;
      let midY;

      if (presetName.includes('3')) {
        midX = x1 + (x2 - x1) * midFraction;
        const pts = `${f1(x1)},${f1(y1)} ${f1(midX)},${f1(y1)} ${f1(midX)},${f1(y2)} ${f1(x2)},${f1(y2)}`;
        sb.appendLine(`<polyline points="${pts}" fill="none" stroke="${strokeColor}" stroke-width="${f2(strokePx)}"${dashAttr}${markerAttrs}/>`);
      } else if (presetName.includes('5')) {
        midX = x1 + (x2 - x1) * midFraction;
        midY = (y1 + y2) / 2.0;
        const pts = `${f1(x1)},${f1(y1)} ${f1(midX)},${f1(y1)} ${f1(midX)},${f1(midY)} ${f1(x2)},${f1(midY)} ${f1(x2)},${f1(y2)}`;
        sb.appendLine(`<polyline points="${pts}" fill="none" stroke="${strokeColor}" stroke-width="${f2(strokePx)}"${dashAttr}${markerAttrs}/>`);
      } else {
        midX = (x1 + x2) / 2.0;
        const pts = `${f1(x1)},${f1(y1)} ${f1(midX)},${f1(y1)} ${f1(midX)},${f1(y2)} ${f1(x2)},${f1(y2)}`;
        sb.appendLine(`<polyline points="${pts}" fill="none" stroke="${strokeColor}" stroke-width="${f2(strokePx)}"${dashAttr}${markerAttrs}/>`);
      }
    } else if (ciIncludes(presetName, 'curvedConnector')) {
      const mcx = (x1 + x2) / 2.0;
      const mcy = (y1 + y2) / 2.0;
      const d = `M${f1(x1)},${f1(y1)} C${f1(mcx)},${f1(y1)} ${f1(mcx)},${f1(y2)} ${f1(x2)},${f1(y2)}`;
      sb.appendLine(`<path d="${d}" fill="none" stroke="${strokeColor}" stroke-width="${f2(strokePx)}"${dashAttr}${markerAttrs}/>`);
    } else {
      sb.appendLine(`<line x1="${f1(x1)}" y1="${f1(y1)}" x2="${f1(x2)}" y2="${f1(y2)}" stroke="${strokeColor}" stroke-width="${f2(strokePx)}"${dashAttr}${markerAttrs}/>`);
    }

    sb.appendLine('</g>');
  },

  buildArrowMarkerDef(id, type, color, strokePx, isStart) {
    const size = Math.max(6, strokePx * 3);
    const orient = isStart ? 'auto-start-reverse' : 'auto';

    if (type === 'triangle') {
      return (
        `<marker id="${id}" markerWidth="${f1(size)}" markerHeight="${f1(size)}" ` +
        `refX="${f1(size)}" refY="${f1(size / 2.0)}" orient="${orient}" markerUnits="userSpaceOnUse">` +
        `<polygon points="0,0 ${f1(size)},${f1(size / 2.0)} 0,${f1(size)}" fill="${color}"/></marker>`
      );
    }
    if (type === 'stealth') {
      return (
        `<marker id="${id}" markerWidth="${f1(size)}" markerHeight="${f1(size)}" ` +
        `refX="${f1(size)}" refY="${f1(size / 2.0)}" orient="${orient}" markerUnits="userSpaceOnUse">` +
        `<polygon points="0,0 ${f1(size)},${f1(size / 2.0)} 0,${f1(size)} ${f1(size * 0.3)},${f1(size / 2.0)}" fill="${color}"/></marker>`
      );
    }
    if (type === 'arrow') {
      return (
        `<marker id="${id}" markerWidth="${f1(size)}" markerHeight="${f1(size)}" ` +
        `refX="${f1(size)}" refY="${f1(size / 2.0)}" orient="${orient}" markerUnits="userSpaceOnUse">` +
        `<polyline points="0,0 ${f1(size)},${f1(size / 2.0)} 0,${f1(size)}" fill="none" stroke="${color}" stroke-width="${f1(Math.max(1, strokePx * 0.5))}"/></marker>`
      );
    }
    if (type === 'diamond') {
      const mid = size / 2.0;
      return (
        `<marker id="${id}" markerWidth="${f1(size)}" markerHeight="${f1(size)}" ` +
        `refX="${f1(mid)}" refY="${f1(mid)}" orient="${orient}" markerUnits="userSpaceOnUse">` +
        `<polygon points="0,${f1(mid)} ${f1(mid)},0 ${f1(size)},${f1(mid)} ${f1(mid)},${f1(size)}" fill="${color}"/></marker>`
      );
    }
    if (type === 'oval') {
      const mid = size / 2.0;
      return (
        `<marker id="${id}" markerWidth="${f1(size)}" markerHeight="${f1(size)}" ` +
        `refX="${f1(mid)}" refY="${f1(mid)}" orient="${orient}" markerUnits="userSpaceOnUse">` +
        `<circle cx="${f1(mid)}" cy="${f1(mid)}" r="${f1(mid)}" fill="${color}"/></marker>`
      );
    }

    return (
      `<marker id="${id}" markerWidth="${f1(size)}" markerHeight="${f1(size)}" ` +
      `refX="${f1(size)}" refY="${f1(size / 2.0)}" orient="${orient}" markerUnits="userSpaceOnUse">` +
      `<polygon points="0,0 ${f1(size)},${f1(size / 2.0)} 0,${f1(size)}" fill="${color}"/></marker>`
    );
  },

  renderGroupShape(sb, defsSb, groupShape, slidePart) {
    const grpSpPr = firstChild(groupShape, 'grpSpPr');
    const xfrm = firstChild(grpSpPr, 'xfrm');
    if (xfrm == null) return;

    const off = firstChild(xfrm, 'off');
    const left = this.emuToPx(attrLong(off, 'x'));
    const top = this.emuToPx(attrLong(off, 'y'));
    const pptxShapeId = this.getPptxShapeId(groupShape) || '';
    const groupDataAttributes = this.buildSvgGroupDataAttributes(pptxShapeId);
    const hiddenStyleAttr = this.buildSvgStyleAttribute(this.buildHiddenSvgStyle(pptxShapeId));

    sb.appendLine(`<g ${groupDataAttributes} transform="translate(${f0(left)},${f0(top)})"${hiddenStyleAttr}>`);

    for (const child of childElements(groupShape)) {
      if (child.localName === 'sp') this.renderShape(sb, defsSb, child, slidePart);
      else if (child.localName === 'pic') this.renderPicture(sb, defsSb, child, slidePart);
      else if (child.localName === 'graphicFrame') this.renderGraphicFrame(sb, defsSb, child, slidePart);
      else if (child.localName === 'cxnSp') this.renderConnectionShape(sb, defsSb, child, slidePart);
    }

    sb.appendLine('</g>');
  },

  renderTable(sb, table) {
    const isTruthyAttr = (el, name) => {
      const v = attr(el, name) || '';
      return v === '1' || ciEq(v, 'true');
    };

    const tblPr = firstChild(table, 'tblPr');
    const firstRowStyled = isTruthyAttr(tblPr, 'firstRow');
    const bandRows = isTruthyAttr(tblPr, 'bandRow');

    const accent1Hex = this.resolveSchemeColor('accent1');
    const dark1Hex = this.resolveSchemeColor('dk1');
    const light1Hex = this.resolveSchemeColor('lt1');
    const headerBg = '#' + accent1Hex;
    const bodyBg = '#FFFFFF';
    const bodyBandBg = '#EEF3F8';
    const borderColor = '#' + dark1Hex;
    const headerTextColor = '#' + light1Hex;
    const bodyTextColor = '#' + dark1Hex;

    sb.appendLine('<table style="width:100%;height:100%;border-collapse:collapse;table-layout:fixed;">');

    const gridCols = elements(firstChild(table, 'tblGrid'), 'gridCol');
    if (gridCols && gridCols.length > 0) {
      const totalW = gridCols.reduce((acc, c) => acc + (attr(c, 'w') != null ? numOr(c, 'w') : 0), 0);
      if (totalW > 0) {
        sb.appendLine('<colgroup>');
        for (const col of gridCols) {
          const wPct = ((attr(col, 'w') != null ? numOr(col, 'w') : 0) * 100.0) / totalW;
          sb.appendLine(`<col style="width:${f3(wPct)}%;"/>`);
        }
        sb.appendLine('</colgroup>');
      }
    }

    let rowIndex = 0;
    for (const row of elements(table, 'tr', A)) {
      const rowHeightPx = this.emuToPx(attr(row, 'h') != null ? numOr(row, 'h') : 0);
      const trStyle = rowHeightPx > 0.1 ? ` style="height:${f1(rowHeightPx)}px;"` : '';
      sb.appendLine(`<tr${trStyle}>`);
      const isHeaderRow = firstRowStyled && rowIndex === 0;
      const isBandRow = bandRows && !isHeaderRow && rowIndex % 2 === 1;

      for (const cell of elements(row, 'tc', A)) {
        let cellStyle = '';
        const tcPr = firstChild(cell, 'tcPr');
        if (tcPr != null) {
          const solidFill = firstChild(tcPr, 'solidFill');
          if (solidFill != null) {
            const color = this.getColorFromFill(solidFill);
            if (color !== '') cellStyle += `background-color:${color};`;
          }
        }

        if (cellStyle === '') {
          if (isHeaderRow) cellStyle += `background-color:${headerBg};color:${headerTextColor};font-weight:700;`;
          else if (isBandRow) cellStyle += `background-color:${bodyBandBg};color:${bodyTextColor};`;
          else cellStyle += `background-color:${bodyBg};color:${bodyTextColor};`;
        }

        cellStyle += `border:1px solid ${borderColor};padding:4px 8px;vertical-align:middle;`;
        sb.append(`<td style="${cellStyle}">`);
        const txBody = textBodyOf(cell);
        if (txBody != null) {
          this.renderTextBody(sb, txBody, this.makeInherited(null, null, null, null, null, null), 1.0, 0.0);
        }
        sb.appendLine('</td>');
      }
      sb.appendLine('</tr>');
      rowIndex++;
    }
    sb.appendLine('</table>');
  },

  // ── Preset Geometry ──────────────────────────────────────────────

  getPresetGeometryData(spPr, widthPx, heightPx) {
    const prstGeom = firstChild(spPr, 'prstGeom');
    if (prstGeom == null) return null;
    const prst = attr(prstGeom, 'prst');
    if (prst == null) return null;

    const avLst = firstChild(prstGeom, 'avLst');

    const readAdjN = (...defaults) => {
      const result = defaults.slice();
      const guides = avLst ? elements(avLst, 'gd') : [];
      for (let i = 0; i < guides.length && i < result.length; i++) {
        const v = parseAdjFormula(attr(guides[i], 'fmla'));
        if (v != null) result[i] = v;
      }
      return result;
    };

    const w = widthPx;
    const h = heightPx;
    const ss = Math.min(w, h);
    const cx = w / 2.0;
    const cy = h / 2.0;

    if (prst === 'rect') return null;

    if (prst === 'roundRect') {
      const adj = readAdjN(16667);
      let r = (ss * adj[0]) / 100000.0;
      r = Math.max(1.0, Math.min(Math.min(w, h) / 2.0, r));
      return { type: 'roundrect', data: '', rx: r, ry: r, inset: geomInset(0, 0, 0, 0) };
    }

    if (prst === 'ellipse')
      return { type: 'ellipse', data: '', rx: cx, ry: cy, inset: geomInset(cy * 0.29, cx * 0.29, cy * 0.29, cx * 0.29) };

    if (prst === 'pie') {
      const adj = readAdjN(0, 16200000);
      const startDeg = normalizeClockwiseAngle(adj[0] / 60000.0);
      const endDeg = normalizeClockwiseAngle(adj[1] / 60000.0);
      let sweepDeg = endDeg - startDeg;
      if (sweepDeg <= 0.0) sweepDeg += 360.0;

      if (sweepDeg >= 359.999)
        return { type: 'ellipse', data: '', rx: cx, ry: cy, inset: geomInset(cy * 0.29, cx * 0.29, cy * 0.29, cx * 0.29) };

      const d = buildPiePath(w, h, startDeg, sweepDeg);
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(cy * 0.29, cx * 0.29, cy * 0.29, cx * 0.29) };
    }

    if (prst === 'triangle') {
      const adj = readAdjN(50000);
      const apexX = (w * adj[0]) / 100000.0;
      const pts = `${f1(apexX)},0 ${f1(w)},${f1(h)} 0,${f1(h)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(h * 0.45, w * 0.12, 0, w * 0.12) };
    }

    if (prst === 'diamond') {
      const pts = `${f1(cx)},0 ${f1(w)},${f1(cy)} ${f1(cx)},${f1(h)} 0,${f1(cy)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(cy * 0.5, cx * 0.5, cy * 0.5, cx * 0.5) };
    }

    if (prst === 'parallelogram') {
      const adj = readAdjN(25000);
      const dx = (ss * adj[0]) / 100000.0;
      const pts = `${f1(dx)},0 ${f1(w)},0 ${f1(w - dx)},${f1(h)} 0,${f1(h)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(0, dx, 0, dx) };
    }

    if (prst === 'trapezoid') {
      const adj = readAdjN(25000);
      const dx = (w * adj[0]) / 100000.0;
      const pts = `${f1(dx)},0 ${f1(w - dx)},0 ${f1(w)},${f1(h)} 0,${f1(h)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(0, dx, 0, dx) };
    }

    if (prst === 'pentagon') {
      const pts = buildRegularPolygonPts(5, w, h);
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(0, 0, 0, 0) };
    }

    if (prst === 'hexagon') {
      const adj = readAdjN(25000);
      const dx = (ss * adj[0]) / 100000.0;
      const pts = `${f1(dx)},0 ${f1(w - dx)},0 ${f1(w)},${f1(cy)} ${f1(w - dx)},${f1(h)} ${f1(dx)},${f1(h)} 0,${f1(cy)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(0, dx, 0, dx) };
    }

    if (prst === 'octagon') {
      const adj = readAdjN(29289);
      const dx = (w * adj[0]) / 100000.0;
      const dy = (h * adj[0]) / 100000.0;
      const pts = `${f1(dx)},0 ${f1(w - dx)},0 ${f1(w)},${f1(dy)} ${f1(w)},${f1(h - dy)} ${f1(w - dx)},${f1(h)} ${f1(dx)},${f1(h)} 0,${f1(h - dy)} 0,${f1(dy)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(dy, dx, dy, dx) };
    }

    if (prst === 'rightArrow') {
      const adj = readAdjN(50000, 50000);
      const dy1 = (h * adj[0]) / 200000.0;
      const y1 = cy - dy1;
      const y2 = cy + dy1;
      const dx1 = (ss * adj[1]) / 100000.0;
      const x1 = w - dx1;
      const pts = `0,${f1(y1)} ${f1(x1)},${f1(y1)} ${f1(x1)},0 ${f1(w)},${f1(cy)} ${f1(x1)},${f1(h)} ${f1(x1)},${f1(y2)} 0,${f1(y2)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(y1, dx1, y1, 0) };
    }
    if (prst === 'leftArrow') {
      const adj = readAdjN(50000, 50000);
      const dy1 = (h * adj[0]) / 200000.0;
      const y1 = cy - dy1;
      const y2 = cy + dy1;
      const dx1 = (ss * adj[1]) / 100000.0;
      const pts = `0,${f1(cy)} ${f1(dx1)},0 ${f1(dx1)},${f1(y1)} ${f1(w)},${f1(y1)} ${f1(w)},${f1(y2)} ${f1(dx1)},${f1(y2)} ${f1(dx1)},${f1(h)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(y1, 0, y1, dx1) };
    }
    if (prst === 'upArrow') {
      const adj = readAdjN(50000, 50000);
      const dx1 = (w * adj[0]) / 200000.0;
      const x1 = cx - dx1;
      const x2 = cx + dx1;
      const dy1 = (ss * adj[1]) / 100000.0;
      const pts = `${f1(cx)},0 ${f1(w)},${f1(dy1)} ${f1(x2)},${f1(dy1)} ${f1(x2)},${f1(h)} ${f1(x1)},${f1(h)} ${f1(x1)},${f1(dy1)} 0,${f1(dy1)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(dy1, x1, 0, x1) };
    }
    if (prst === 'downArrow') {
      const adj = readAdjN(50000, 50000);
      const dx1 = (w * adj[0]) / 200000.0;
      const x1 = cx - dx1;
      const x2 = cx + dx1;
      const dy1 = (ss * adj[1]) / 100000.0;
      const y1 = h - dy1;
      const pts = `${f1(x1)},0 ${f1(x2)},0 ${f1(x2)},${f1(y1)} ${f1(w)},${f1(y1)} ${f1(cx)},${f1(h)} 0,${f1(y1)} ${f1(x1)},${f1(y1)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(0, x1, dy1, x1) };
    }
    if (prst === 'leftRightArrow') {
      const adj = readAdjN(50000, 50000);
      const dy1 = (h * adj[0]) / 200000.0;
      const y1 = cy - dy1;
      const y2 = cy + dy1;
      const dx1 = (ss * adj[1]) / 100000.0;
      const pts = `0,${f1(cy)} ${f1(dx1)},0 ${f1(dx1)},${f1(y1)} ${f1(w - dx1)},${f1(y1)} ${f1(w - dx1)},0 ${f1(w)},${f1(cy)} ${f1(w - dx1)},${f1(h)} ${f1(w - dx1)},${f1(y2)} ${f1(dx1)},${f1(y2)} ${f1(dx1)},${f1(h)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(y1, dx1, y1, dx1) };
    }
    if (prst === 'upDownArrow') {
      const adj = readAdjN(50000, 50000);
      const dx1 = (w * adj[0]) / 200000.0;
      const x1 = cx - dx1;
      const x2 = cx + dx1;
      const dy1 = (ss * adj[1]) / 100000.0;
      const pts = `${f1(cx)},0 ${f1(w)},${f1(dy1)} ${f1(x2)},${f1(dy1)} ${f1(x2)},${f1(h - dy1)} ${f1(w)},${f1(h - dy1)} ${f1(cx)},${f1(h)} 0,${f1(h - dy1)} ${f1(x1)},${f1(h - dy1)} ${f1(x1)},${f1(dy1)} 0,${f1(dy1)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(dy1, x1, dy1, x1) };
    }
    if (prst === 'notchedRightArrow') {
      const adj = readAdjN(50000, 50000);
      const dy1 = (h * adj[0]) / 200000.0;
      const y1 = cy - dy1;
      const y2 = cy + dy1;
      const dx1 = (ss * adj[1]) / 100000.0;
      const x1 = w - dx1;
      const notch = dx1 * 0.5;
      const pts = `0,${f1(y1)} ${f1(x1)},${f1(y1)} ${f1(x1)},0 ${f1(w)},${f1(cy)} ${f1(x1)},${f1(h)} ${f1(x1)},${f1(y2)} 0,${f1(y2)} ${f1(notch)},${f1(cy)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(y1, dx1, y1, 0) };
    }

    if (prst === 'chevron') {
      const adj = readAdjN(50000);
      const x1 = (ss * adj[0]) / 100000.0;
      const x2 = w - x1;
      const pts = `0,0 ${f1(x2)},0 ${f1(w)},${f1(cy)} ${f1(x2)},${f1(h)} 0,${f1(h)} ${f1(x1)},${f1(cy)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(0, x1, 0, x1) };
    }

    if (prst === 'homePlate') {
      const adj = readAdjN(50000);
      const dx = (ss * adj[0]) / 100000.0;
      const pts = `0,0 ${f1(w - dx)},0 ${f1(w)},${f1(cy)} ${f1(w - dx)},${f1(h)} 0,${f1(h)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(0, dx * 0.5, 0, 0) };
    }

    const star = isStarPreset(prst);
    if (star != null) {
      const adj = readAdjN(star.defaultAdj);
      const innerFraction = adj[0] / 50000.0;
      const pts = buildStarPolygonPts(star.points, w, h, innerFraction);
      const inscribed = innerFraction * Math.cos(Math.PI / star.points);
      const starInsetX = cx * (1.0 - inscribed);
      const starInsetY = cy * (1.0 - inscribed);
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(starInsetY, starInsetX, starInsetY, starInsetX) };
    }

    if (prst === 'irregularSeal1') {
      const d = `M${f1(0.15 * w)},${f1(0.0 * h)} L${f1(0.55 * w)},${f1(0.13 * h)} L${f1(0.48 * w)},${f1(0.0 * h)} L${f1(0.72 * w)},${f1(0.22 * h)} L${f1(w)},${f1(0.15 * h)} L${f1(0.75 * w)},${f1(0.45 * h)} L${f1(w)},${f1(0.55 * h)} L${f1(0.72 * w)},${f1(0.62 * h)} L${f1(0.88 * w)},${f1(h)} L${f1(0.55 * w)},${f1(0.72 * h)} L${f1(0.35 * w)},${f1(h)} L${f1(0.38 * w)},${f1(0.68 * h)} L${f1(0.0 * w)},${f1(0.72 * h)} L${f1(0.25 * w)},${f1(0.5 * h)} L${f1(0.0 * w)},${f1(0.38 * h)} L${f1(0.2 * w)},${f1(0.32 * h)} Z`;
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(0, 0, 0, 0) };
    }
    if (prst === 'irregularSeal2') {
      const d = `M${f1(0.12 * w)},${f1(0.0 * h)} L${f1(0.38 * w)},${f1(0.18 * h)} L${f1(0.58 * w)},${f1(0.02 * h)} L${f1(0.55 * w)},${f1(0.25 * h)} L${f1(0.82 * w)},${f1(0.12 * h)} L${f1(0.72 * w)},${f1(0.35 * h)} L${f1(w)},${f1(0.38 * h)} L${f1(0.75 * w)},${f1(0.52 * h)} L${f1(0.95 * w)},${f1(0.68 * h)} L${f1(0.68 * w)},${f1(0.62 * h)} L${f1(0.72 * w)},${f1(0.88 * h)} L${f1(0.5 * w)},${f1(0.72 * h)} L${f1(0.35 * w)},${f1(h)} L${f1(0.32 * w)},${f1(0.7 * h)} L${f1(0.08 * w)},${f1(0.82 * h)} L${f1(0.18 * w)},${f1(0.58 * h)} L${f1(0.0 * w)},${f1(0.5 * h)} L${f1(0.18 * w)},${f1(0.38 * h)} L${f1(0.05 * w)},${f1(0.22 * h)} L${f1(0.22 * w)},${f1(0.22 * h)} Z`;
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(0, 0, 0, 0) };
    }

    if (prst === 'wedgeRoundRectCallout') {
      const adj = readAdjN(-20833, 62500, 16667);
      const ptrX = cx + (w * adj[0]) / 100000.0;
      const ptrY = cy + (h * adj[1]) / 100000.0;
      let r = (ss * adj[2]) / 100000.0;
      r = Math.max(1.0, Math.min(Math.min(w, h) / 2.0, r));
      const d = buildWedgeRoundRectPath(w, h, r, ptrX, ptrY);
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(0, 0, 0, 0) };
    }
    if (prst === 'wedgeRectCallout') {
      const adj = readAdjN(-20833, 62500);
      const ptrX = cx + (w * adj[0]) / 100000.0;
      const ptrY = cy + (h * adj[1]) / 100000.0;
      const d = buildWedgeRectPath(w, h, ptrX, ptrY);
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(0, 0, 0, 0) };
    }
    if (prst === 'wedgeEllipseCallout') {
      const rx = cx;
      const ry = cy;
      const adj = readAdjN(-20833, 62500);
      const ptrX = cx + (w * adj[0]) / 100000.0;
      const ptrY = cy + (h * adj[1]) / 100000.0;
      const angle = Math.atan2(ptrY - cy, ptrX - cx);
      const tipW = 0.1;
      const bx1 = cx + rx * Math.cos(angle - tipW);
      const by1 = cy + ry * Math.sin(angle - tipW);
      const bx2 = cx + rx * Math.cos(angle + tipW);
      const by2 = cy + ry * Math.sin(angle + tipW);
      const d = `M${f1(cx + rx)},${f1(cy)} A${f1(rx)},${f1(ry)} 0 0 1 ${f1(bx1)},${f1(by1)} L${f1(ptrX)},${f1(ptrY)} L${f1(bx2)},${f1(by2)} A${f1(rx)},${f1(ry)} 0 1 1 ${f1(cx + rx)},${f1(cy)} Z`;
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(cy * 0.29, cx * 0.29, cy * 0.29, cx * 0.29) };
    }

    if (prst === 'plus') {
      const adj = readAdjN(25000);
      const dx = (w * adj[0]) / 100000.0;
      const dy = (h * adj[0]) / 100000.0;
      const pts = `${f1(dx)},0 ${f1(w - dx)},0 ${f1(w - dx)},${f1(dy)} ${f1(w)},${f1(dy)} ${f1(w)},${f1(h - dy)} ${f1(w - dx)},${f1(h - dy)} ${f1(w - dx)},${f1(h)} ${f1(dx)},${f1(h)} ${f1(dx)},${f1(h - dy)} 0,${f1(h - dy)} 0,${f1(dy)} ${f1(dx)},${f1(dy)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(dy, dx, dy, dx) };
    }

    if (prst === 'rtTriangle') {
      const pts = `0,0 ${f1(w)},${f1(h)} 0,${f1(h)}`;
      return { type: 'polygon', data: pts, rx: 0, ry: 0, inset: geomInset(0, w * 0.5, 0, 0) };
    }

    if (prst === 'stripedRightArrow') {
      const adj = readAdjN(50000, 50000);
      const dy1 = (h * adj[0]) / 200000.0;
      const y1 = cy - dy1;
      const y2 = cy + dy1;
      const dx1 = (ss * adj[1]) / 100000.0;
      const x1 = w - dx1;
      const stripe = w * 0.05;
      const d = `M0,${f1(y1)} L${f1(stripe * 0.4)},${f1(y1)} L${f1(stripe * 0.4)},${f1(y2)} L0,${f1(y2)} Z M${f1(stripe * 0.7)},${f1(y1)} L${f1(stripe)},${f1(y1)} L${f1(stripe)},${f1(y2)} L${f1(stripe * 0.7)},${f1(y2)} Z M${f1(stripe * 1.3)},${f1(y1)} L${f1(x1)},${f1(y1)} L${f1(x1)},0 L${f1(w)},${f1(cy)} L${f1(x1)},${f1(h)} L${f1(x1)},${f1(y2)} L${f1(stripe * 1.3)},${f1(y2)} Z`;
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(y1, dx1, y1, 0) };
    }

    if (prst === 'cloud') {
      const d =
        `M${f1(0.146 * w)},${f1(0.812 * h)} ` +
        `C${f1(0.024 * w)},${f1(0.742 * h)} ${f1(0.0 * w)},${f1(0.588 * h)} ${f1(0.064 * w)},${f1(0.472 * h)} ` +
        `C${f1(0.018 * w)},${f1(0.37 * h)} ${f1(0.046 * w)},${f1(0.242 * h)} ${f1(0.154 * w)},${f1(0.178 * h)} ` +
        `C${f1(0.206 * w)},${f1(0.068 * h)} ${f1(0.35 * w)},${f1(0.0 * h)} ${f1(0.5 * w)},${f1(0.032 * h)} ` +
        `C${f1(0.622 * w)},${f1(0.0 * h)} ${f1(0.748 * w)},${f1(0.042 * h)} ${f1(0.81 * w)},${f1(0.15 * h)} ` +
        `C${f1(0.93 * w)},${f1(0.17 * h)} ${f1(w)},${f1(0.32 * h)} ${f1(0.976 * w)},${f1(0.462 * h)} ` +
        `C${f1(w)},${f1(0.596 * h)} ${f1(0.952 * w)},${f1(0.72 * h)} ${f1(0.858 * w)},${f1(0.79 * h)} ` +
        `C${f1(0.82 * w)},${f1(0.92 * h)} ${f1(0.7 * w)},${f1(h)} ${f1(0.558 * w)},${f1(0.978 * h)} ` +
        `C${f1(0.42 * w)},${f1(h)} ${f1(0.29 * w)},${f1(0.94 * h)} ${f1(0.244 * w)},${f1(0.862 * h)} ` +
        `C${f1(0.196 * w)},${f1(0.876 * h)} ${f1(0.168 * w)},${f1(0.856 * h)} ${f1(0.146 * w)},${f1(0.812 * h)} Z`;
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(h * 0.15, w * 0.1, h * 0.1, w * 0.15) };
    }

    if (prst === 'heart') {
      const d =
        `M${f1(cx)},${f1(h)} ` +
        `C${f1(0.0)},${f1(0.6 * h)} ${f1(0.0)},${f1(0.0)} ${f1(cx)},${f1(0.3 * h)} ` +
        `C${f1(w)},${f1(0.0)} ${f1(w)},${f1(0.6 * h)} ${f1(cx)},${f1(h)} Z`;
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(0, 0, 0, 0) };
    }

    if (prst === 'lightningBolt') {
      const d = `M${f1(0.38 * w)},0 L${f1(0.56 * w)},${f1(0.35 * h)} L${f1(0.78 * w)},${f1(0.25 * h)} L${f1(0.5 * w)},${f1(0.6 * h)} L${f1(w)},${f1(0.52 * h)} L${f1(0.32 * w)},${f1(h)} L${f1(0.42 * w)},${f1(0.62 * h)} L${f1(0.18 * w)},${f1(0.7 * h)} L${f1(0.44 * w)},${f1(0.38 * h)} L0,${f1(0.48 * h)} Z`;
      return { type: 'path', data: d, rx: 0, ry: 0, inset: geomInset(0, 0, 0, 0) };
    }

    return null;
  },

  // ── Custom Geometry ──────────────────────────────────────────────

  getCustomGeometrySvgElements(spPr, widthPx, heightPx, fillAttr, defsSb) {
    const customGeom = firstChild(spPr, 'custGeom');
    const pathList = firstChild(customGeom, 'pathLst');
    const paths = pathList ? elements(pathList, 'path', A) : [];
    if (paths.length === 0) return { fillElement: '', strokeElement: '', clipId: '' };

    // NOTE: The C# code intends to "prefer a non-fill=none path", but with
    // DocumentFormat.OpenXml 3.4.1 `Path.Fill.Value.ToString()` returns the
    // record-struct default ("PathFillModeValues { }"), never the token
    // "none" — so its `!fillMode.Equals("none")` guard is always true and the
    // loop always breaks on the first path. Replicate that exactly.
    let clipPath = paths[0];

    let strokePath = paths[0];
    for (const p of paths) {
      const strokeDisabled = attr(p, 'stroke') != null && !(attr(p, 'stroke') === '1' || ciEq(attr(p, 'stroke'), 'true'));
      if (!strokeDisabled) {
        strokePath = p;
        break;
      }
    }

    const clipD = this.buildSvgPathData(clipPath, widthPx, heightPx);
    if (clipD === '') return { fillElement: '', strokeElement: '', clipId: '' };

    const fillElement = `<path d="${clipD}" ${fillAttr}/>`;

    const clipId = `clip${++this._defsCounter}`;
    defsSb.appendLine(`<clipPath id="${clipId}"><path d="${clipD}"/></clipPath>`);

    let strokeElement = '';
    const outline = firstChild(spPr, 'ln');
    if (outline != null && firstChild(outline, 'noFill') == null) {
      const strokeD = this.buildSvgPathData(strokePath, widthPx, heightPx);
      if (strokeD !== '') {
        const outlineWidthEmu = attr(outline, 'w') != null ? numOr(outline, 'w') : 12700;
        let strokePx = this.emuToPx(outlineWidthEmu);
        if (strokePx <= 0) strokePx = 1;

        let strokeColor = '#000000';
        const solidFill = firstChild(outline, 'solidFill');
        if (solidFill != null) {
          const c = this.getColorFromFill(solidFill);
          if (c !== '') strokeColor = c;
        }

        let lineCap = 'butt';
        const capAttr = attr(outline, 'cap') || '';
        if (ciEq(capAttr, 'rnd')) lineCap = 'round';
        else if (ciEq(capAttr, 'sq')) lineCap = 'square';

        let dashAttr = '';
        const dash = firstChild(outline, 'prstDash');
        if (dash != null && attr(dash, 'val') != null) {
          const v = attr(dash, 'val') || '';
          const a = Math.max(1, strokePx);
          let dashArray = null;
          if (ciIncludes(v, 'dot')) {
            if (ciIncludes(v, 'DashDotDot')) dashArray = `${f1(4 * a)},${f1(2 * a)},${f1(a)},${f1(2 * a)},${f1(a)},${f1(2 * a)}`;
            else if (ciIncludes(v, 'dashDot')) dashArray = `${f1(4 * a)},${f1(2 * a)},${f1(a)},${f1(2 * a)}`;
            else dashArray = `${f1(a)},${f1(2 * a)}`;
          } else if (!ciEq(v, 'solid')) {
            const isLongDash = ciIncludes(v, 'lg');
            dashArray = isLongDash ? `${f1(6 * a)},${f1(3 * a)}` : `${f1(4 * a)},${f1(3 * a)}`;
          }
          if (dashArray) dashAttr = ` stroke-dasharray="${dashArray}"`;
        }

        strokeElement = `<path d="${strokeD}" fill="none" stroke="${strokeColor}" stroke-width="${f2(strokePx)}" stroke-linecap="${lineCap}" stroke-linejoin="round"${dashAttr}/>`;
      }
    }

    return { fillElement, strokeElement, clipId };
  },

  buildSvgPathData(path, outWidthPx, outHeightPx) {
    const srcW = attr(path, 'w') != null ? numOr(path, 'w') : outWidthPx;
    const srcH = attr(path, 'h') != null ? numOr(path, 'h') : outHeightPx;
    if (srcW === 0 || srcH === 0) return '';
    const sx = outWidthPx / srcW;
    const sy = outHeightPx / srcH;

    const toX = (pt) => {
      const xStr = attr(pt, 'x') != null ? attr(pt, 'x') : '0';
      const x = parseInt(xStr, 10);
      return Number.isNaN(x) ? 0 : x * sx;
    };
    const toY = (pt) => {
      const yStr = attr(pt, 'y') != null ? attr(pt, 'y') : '0';
      const y = parseInt(yStr, 10);
      return Number.isNaN(y) ? 0 : y * sy;
    };

    let d = '';
    for (const seg of childElements(path)) {
      if (seg.localName === 'moveTo') {
        const pt = firstChild(seg, 'pt');
        if (pt != null) d += `M${f1(toX(pt))} ${f1(toY(pt))} `;
      } else if (seg.localName === 'lnTo') {
        const pt = firstChild(seg, 'pt');
        if (pt != null) d += `L${f1(toX(pt))} ${f1(toY(pt))} `;
      } else if (seg.localName === 'cubicBezTo') {
        const pts = elements(seg, 'pt', A).slice(0, 3);
        if (pts.length === 3) {
          d += `C${f1(toX(pts[0]))} ${f1(toY(pts[0]))} ${f1(toX(pts[1]))} ${f1(toY(pts[1]))} ${f1(toX(pts[2]))} ${f1(toY(pts[2]))} `;
        }
      } else if (seg.localName === 'quadBezTo') {
        const pts = elements(seg, 'pt', A).slice(0, 2);
        if (pts.length === 2) {
          d += `Q${f1(toX(pts[0]))} ${f1(toY(pts[0]))} ${f1(toX(pts[1]))} ${f1(toY(pts[1]))} `;
        }
      } else if (seg.localName === 'arcTo') {
        const parseL = (name) => {
          const v = attr(seg, name);
          const n = parseInt(v != null ? v : '0', 10);
          return Number.isNaN(n) ? 0 : n;
        };
        const wR = parseL('wR') * sx;
        const hR = parseL('hR') * sy;
        const stAng = parseL('stAng') / 60000.0;
        const swAng = parseL('swAng') / 60000.0;

        if (wR > 0 && hR > 0 && Math.abs(swAng) > 0.001) {
          const endRad = ((stAng + swAng) * Math.PI) / 180.0;
          const endX = wR * Math.cos(endRad);
          const endY = hR * Math.sin(endRad);
          const largeArc = Math.abs(swAng) > 180 ? 1 : 0;
          const sweepFlag = swAng > 0 ? 1 : 0;
          d += `A${f1(wR)} ${f1(hR)} 0 ${largeArc} ${sweepFlag} ${f1(endX)} ${f1(endY)} `;
        }
      } else if (seg.localName === 'close') {
        d += 'Z ';
      }
    }
    return d.trim();
  },

  // ── Transform / Position helpers ─────────────────────────────────

  getTransform(spPr) {
    const xfrm = firstChild(spPr, 'xfrm');
    if (xfrm == null) return [0, 0, 0, 0];

    const off = firstChild(xfrm, 'off');
    const ext = firstChild(xfrm, 'ext');
    const left = this.emuToPx(attrLong(off, 'x'));
    const top = this.emuToPx(attrLong(off, 'y'));
    const width = this.emuToPx(attrLong(ext, 'cx'));
    const height = this.emuToPx(attrLong(ext, 'cy'));

    return [left, top, width, height];
  },

  getScene3DTransform(spPr) {
    if (spPr == null) return '';

    const scene3d = childElements(spPr).find((e) => e.localName === 'scene3d');
    const camera = scene3d ? childElements(scene3d).find((e) => e.localName === 'camera') : null;
    if (camera == null) return '';

    const preset = attr(camera, 'prst') || '';
    if (preset === '') return '';

    if (ciIncludes(preset, 'LeftFacing')) return 'perspective(1200px) rotateY(-18deg) rotateX(4deg)';
    if (ciIncludes(preset, 'RightFacing')) return 'perspective(1200px) rotateY(18deg) rotateX(4deg)';
    if (ciIncludes(preset, 'Relaxed')) return 'perspective(1400px) rotateX(2deg)';
    if (ciIncludes(preset, 'Perspective')) return 'perspective(1200px)';

    return '';
  },

  // ── Effect color helpers ─────────────────────────────────────────

  getColorFromEffectColorContainer(colorContainer, fallbackHexRgbOrRgba) {
    if (colorContainer == null) return cssColorFromHexWithAlphaRaw(fallbackHexRgbOrRgba);

    const srgb = firstChild(colorContainer, 'srgbClr');
    if (srgb != null && attr(srgb, 'val') != null)
      return cssColorFromHexWithAlphaRaw(this.applyColorTransforms(attr(srgb, 'val'), childElements(srgb)));

    const scheme = firstChild(colorContainer, 'schemeClr');
    if (scheme != null && attr(scheme, 'val') != null) {
      const hex = this.resolveSchemeColor(attr(scheme, 'val'));
      return cssColorFromHexWithAlphaRaw(this.applyColorTransforms(hex, childElements(scheme)));
    }

    const preset = firstChild(colorContainer, 'prstClr');
    if (preset != null && attr(preset, 'val') != null) {
      const { PRESET_COLORS } = require('./colorHelper');
      if (PRESET_COLORS.has(attr(preset, 'val'))) {
        const mapped = PRESET_COLORS.get(attr(preset, 'val'));
        const rgb = mapped.startsWith('#') ? mapped.slice(1) : mapped;
        return cssColorFromHexWithAlphaRaw(this.applyColorTransforms(rgb, childElements(preset)));
      }
    }

    return cssColorFromHexWithAlphaRaw(fallbackHexRgbOrRgba);
  },

  // ── Placeholder resolution ───────────────────────────────────────

  resolvePhTransform(ph, slidePart) {
    const { slideLayoutPart, slideMasterPart } = require('../ooxml/parts');
    const layoutPart = slideLayoutPart(slidePart);
    if (layoutPart != null) {
      let match = this.findMatchingPlaceholder(ph, shapeTreeOf(layoutPart.root));
      if (match != null && shapePropsOf(match) != null) {
        const t = this.getTransform(shapePropsOf(match));
        if (t[2] > 0 && t[3] > 0) return t;
      }

      const masterPart = slideMasterPart(layoutPart);
      if (masterPart != null) {
        match = this.findMatchingPlaceholder(ph, shapeTreeOf(masterPart.root));
        if (match != null && shapePropsOf(match) != null) {
          const t = this.getTransform(shapePropsOf(match));
          if (t[2] > 0 && t[3] > 0) return t;
        }
      }
    }
    return null;
  },

  resolveStyleFillRef(fillRef) {
    if (fillRef == null || attr(fillRef, 'idx') == null || numOr(fillRef, 'idx') === 0) return 'fill="none"';

    const hex = this.resolveStyleRefColor(fillRef);
    if (hex === '') return 'fill="none"';
    return this.svgFillAttr('#' + hex);
  },

  resolveStyleLineRef(lnRef) {
    if (lnRef == null || attr(lnRef, 'idx') == null || numOr(lnRef, 'idx') === 0) return 'stroke="none"';

    const hex = this.resolveStyleRefColor(lnRef);
    if (hex === '') return 'stroke="none"';

    const idx = Math.trunc(numOr(lnRef, 'idx'));
    let widthEmu = 12700;
    if (idx >= 1 && idx <= this._themeLineWidths.length) widthEmu = this._themeLineWidths[idx - 1];
    const strokePx = this.emuToPx(widthEmu);

    return `stroke="#${hex}" stroke-width="${f2(strokePx)}"`;
  },

  resolveStyleRefColor(styleRef) {
    const schemeClr = firstChild(styleRef, 'schemeClr');
    if (schemeClr != null && attr(schemeClr, 'val') != null) {
      const hex = this.resolveSchemeColor(attr(schemeClr, 'val'));
      return this.applyColorTransforms(hex, childElements(schemeClr));
    }

    const srgb = firstChild(styleRef, 'srgbClr');
    if (srgb != null && attr(srgb, 'val') != null) return this.applyColorTransforms(attr(srgb, 'val'), childElements(srgb));

    return '';
  },

  findMatchingPlaceholder(ph, shapeTree) {
    if (shapeTree == null) return null;

    const phType = attr(ph, 'type');
    const phIdx = attr(ph, 'idx');

    let typeMatch = null;

    for (const sp of elements(shapeTree, 'sp', P)) {
      const spPh = placeholderOf(sp);
      if (spPh == null) continue;

      const spType = attr(spPh, 'type');
      const spIdx = attr(spPh, 'idx');

      if (phIdx != null && spIdx != null && phIdx === spIdx) return sp;

      if (phIdx == null && spIdx == null && phType != null && spType != null && phType === spType) return sp;

      if (phType != null && spType != null && phType === spType) typeMatch = sp;
    }

    return typeMatch;
  },
};

// ── module-scope pure helpers ──────────────────────────────────────────────

function isStarPreset(prst) {
  switch (prst) {
    case 'star4': return { points: 4, defaultAdj: 12500 };
    case 'star5': return { points: 5, defaultAdj: 19098 };
    case 'star6': return { points: 6, defaultAdj: 28868 };
    case 'star8': return { points: 8, defaultAdj: 37500 };
    case 'star10': return { points: 10, defaultAdj: 42533 };
    case 'star12': return { points: 12, defaultAdj: 37500 };
    case 'star16': return { points: 16, defaultAdj: 37500 };
    case 'star24': return { points: 24, defaultAdj: 37500 };
    case 'star32': return { points: 32, defaultAdj: 37500 };
    default: return null;
  }
}

function buildStarPolygonPts(n, w, h, innerFraction) {
  const cx = w / 2.0;
  const cy = h / 2.0;
  const parts = [];
  for (let i = 0; i < 2 * n; i++) {
    const angle = (i * Math.PI) / n - Math.PI / 2.0;
    const isOuter = i % 2 === 0;
    const r = isOuter ? 1.0 : innerFraction;
    const px = cx + cx * r * Math.cos(angle);
    const py = cy + cy * r * Math.sin(angle);
    parts.push(`${f1(px)},${f1(py)}`);
  }
  return parts.join(' ');
}

function buildRegularPolygonPts(n, w, h) {
  const cx = w / 2.0;
  const cy = h / 2.0;
  const parts = [];
  for (let i = 0; i < n; i++) {
    const angle = (i * 2 * Math.PI) / n - Math.PI / 2.0;
    const px = cx + cx * Math.cos(angle);
    const py = cy + cy * Math.sin(angle);
    parts.push(`${f1(px)},${f1(py)}`);
  }
  return parts.join(' ');
}

function normalizeClockwiseAngle(angleDeg) {
  const normalized = angleDeg % 360.0;
  return normalized < 0 ? normalized + 360.0 : normalized;
}

function buildPiePath(w, h, startDeg, sweepDeg) {
  const cx = w / 2.0;
  const cy = h / 2.0;
  const rx = w / 2.0;
  const ry = h / 2.0;
  const startRad = (startDeg * Math.PI) / 180.0;
  const endRad = ((startDeg + sweepDeg) * Math.PI) / 180.0;
  const startX = cx + rx * Math.cos(startRad);
  const startY = cy + ry * Math.sin(startRad);
  const endX = cx + rx * Math.cos(endRad);
  const endY = cy + ry * Math.sin(endRad);
  const largeArc = sweepDeg > 180.0 ? 1 : 0;

  return `M${f1(startX)},${f1(startY)} A${f1(rx)},${f1(ry)} 0 ${largeArc} 1 ${f1(endX)},${f1(endY)} L${f1(cx)},${f1(cy)} Z`;
}

function buildWedgeRoundRectPath(w, h, r, ptrX, ptrY) {
  const cx = w / 2.0;
  const cy = h / 2.0;
  const dxPos = ptrX - cx;
  const dyPos = ptrY - cy;

  const dq = w > 0 ? (dxPos * h) / w : 0;
  const dz = Math.abs(dyPos) - Math.abs(dq);

  const x1 = (w * (dxPos >= 0 ? 7.0 : 2.0)) / 12.0;
  const x2 = (w * (dxPos >= 0 ? 10.0 : 5.0)) / 12.0;
  const y1 = (h * (dyPos >= 0 ? 7.0 : 2.0)) / 12.0;
  const y2 = (h * (dyPos >= 0 ? 10.0 : 5.0)) / 12.0;

  const xt = dz >= 0 ? (dyPos >= 0 ? x1 : ptrX) : x1;
  const yt = dz >= 0 ? (dyPos >= 0 ? 0 : ptrY) : 0;
  const xr = dz >= 0 ? w : dxPos >= 0 ? ptrX : w;
  const yr = dz >= 0 ? y1 : dxPos >= 0 ? ptrY : y1;
  const xb = dz >= 0 ? (dyPos >= 0 ? ptrX : x1) : x1;
  const yb = dz >= 0 ? (dyPos >= 0 ? ptrY : h) : h;
  const xl = dz >= 0 ? 0 : dxPos >= 0 ? 0 : ptrX;
  const yl = dz >= 0 ? y1 : dxPos >= 0 ? y1 : ptrY;

  let d = '';
  d += `M${f1(r)},0 `;
  d += `L${f1(x1)},0 L${f1(xt)},${f1(yt)} L${f1(x2)},0 `;
  d += `L${f1(w - r)},0 Q${f1(w)},0 ${f1(w)},${f1(r)} `;
  d += `L${f1(w)},${f1(y1)} L${f1(xr)},${f1(yr)} L${f1(w)},${f1(y2)} `;
  d += `L${f1(w)},${f1(h - r)} Q${f1(w)},${f1(h)} ${f1(w - r)},${f1(h)} `;
  d += `L${f1(x2)},${f1(h)} L${f1(xb)},${f1(yb)} L${f1(x1)},${f1(h)} `;
  d += `L${f1(r)},${f1(h)} Q0,${f1(h)} 0,${f1(h - r)} `;
  d += `L0,${f1(y2)} L${f1(xl)},${f1(yl)} L0,${f1(y1)} `;
  d += `L0,${f1(r)} Q0,0 ${f1(r)},0 Z`;
  return d;
}

function buildWedgeRectPath(w, h, ptrX, ptrY) {
  return buildWedgeRoundRectPath(w, h, 0, ptrX, ptrY);
}

function cssColorFromHexWithAlphaRaw(hex) {
  const h = (hex || '').trim().replace(/^#/, '');
  if (h.length < 6) return 'rgba(0,0,0,1)';

  const r = parseInt(h.substr(0, 2), 16);
  const g = parseInt(h.substr(2, 2), 16);
  const b = parseInt(h.substr(4, 2), 16);
  const a = h.length >= 8 ? parseInt(h.substr(6, 2), 16) / 255.0 : 1.0;

  return `rgba(${r},${g},${b},${f3(a)})`;
}
