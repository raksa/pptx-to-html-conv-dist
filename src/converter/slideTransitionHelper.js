'use strict';

// Port of PptxToHtml/pptx/SlideTransitionHelper.cs

const crypto = require('crypto');
const { DOMParser } = require('@xmldom/xmldom');
const { firstChild, elements, descendants, attr, childElements, innerText, outerXml, NS } = require('../ooxml/dom');
const { StringBuilder } = require('../util/stringbuilder');
const { fmtTrim, bankRound } = require('../util/format');
const { textBodyOf, shapeTreeOf, shapePropsOf } = require('./elements');

const _svgParser = new DOMParser({ onError: () => {} });

const EMPTY_TRANSITION = { type: '', direction: '', orientation: '', pattern: '', durationMs: 0 };

function sha256Hex(buf) {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

function eqIgnoreCase(a, b) {
  return (a || '').toLowerCase() === (b || '').toLowerCase();
}

function normalizeDirectionalTransitionValue(value) {
  switch ((value || '').trim().toLowerCase()) {
    case 'l':
    case 'left': return 'left';
    case 'r':
    case 'right': return 'right';
    case 'u':
    case 'up': return 'up';
    case 'd':
    case 'down': return 'down';
    default: return '';
  }
}

function normalizeTransitionOrientation(value) {
  switch ((value || '').trim().toLowerCase()) {
    case 'horz':
    case 'horizontal': return 'horizontal';
    case 'vert':
    case 'vertical': return 'vertical';
    default: return '';
  }
}

function normalizeSplitMode(value) {
  switch ((value || '').trim().toLowerCase()) {
    case 'in': return 'in';
    case 'out': return 'out';
    default: return '';
  }
}

function getTransitionTranslateOffsets(direction, defaultDirection) {
  const normalizedDirection = !direction || direction.trim() === '' ? defaultDirection : direction;
  switch (normalizedDirection) {
    case 'left': return ['100%', '0%'];
    case 'right': return ['-100%', '0%'];
    case 'up': return ['0%', '100%'];
    case 'down': return ['0%', '-100%'];
    default: return ['100%', '0%'];
  }
}

function resolveTransitionClipEdge(direction, defaultEdge) {
  switch (direction) {
    case 'left': return 'right';
    case 'right': return 'left';
    case 'up': return 'bottom';
    case 'down': return 'top';
    default: return defaultEdge;
  }
}

function getDirectionalClipAnimationName(direction, defaultEdge) {
  const edge = resolveTransitionClipEdge(direction, defaultEdge);
  switch (edge) {
    case 'right': return 'pptxClipInFromRight';
    case 'top': return 'pptxClipInFromTop';
    case 'bottom': return 'pptxClipInFromBottom';
    default: return 'pptxClipInFromLeft';
  }
}

function getSplitClipAnimationName(orientation) {
  return eqIgnoreCase(orientation, 'horizontal') ? 'pptxSplitInHorizontal' : 'pptxSplitInVertical';
}

function getShapeClipAnimationName(pattern) {
  return eqIgnoreCase(pattern, 'circle') ? 'pptxCircleIn' : 'pptxFadeIn';
}

function getTransitionSpeedDurationMs(speed) {
  switch ((speed || '').trim().toLowerCase()) {
    case 'slow': return 2000;
    case 'medium': return 1000;
    case 'med': return 1000;
    case 'fast': return 500;
    default: return 1000;
  }
}

module.exports = {
  extractSlideTransitionInfo(slidePart) {
    const root = slidePart.root;
    if (root == null) return { ...EMPTY_TRANSITION };

    const transitionElements = descendants(root, 'transition', NS.P);
    if (transitionElements.length === 0) return { ...EMPTY_TRANSITION };

    let explicitDurationMs = 0;
    for (const t of transitionElements) {
      const d = this._getExplicitTransitionDurationMs(t);
      if (d > 0) { explicitDurationMs = d; break; }
    }

    for (const transitionElement of transitionElements) {
      const supportedTransition = this._getSupportedSlideTransitionInfo(transitionElement);
      if (supportedTransition == null) continue;

      let durationMs = this._getExplicitTransitionDurationMs(transitionElement);
      if (durationMs <= 0) {
        durationMs =
          explicitDurationMs > 0
            ? explicitDurationMs
            : getTransitionSpeedDurationMs(attr(transitionElement, 'spd'));
      }

      return {
        type: supportedTransition.type,
        direction: supportedTransition.direction || '',
        orientation: supportedTransition.orientation || '',
        pattern: supportedTransition.pattern || '',
        durationMs,
      };
    }

    return { ...EMPTY_TRANSITION };
  },

  buildSlideTransitionStyle(slideTransitionInfo) {
    if (slideTransitionInfo == null || !slideTransitionInfo.type || slideTransitionInfo.type.trim() === '') return '';

    const duration = fmtTrim(this.toCssDurationSeconds(slideTransitionInfo.durationMs, 0.5), 3);

    switch (slideTransitionInfo.type) {
      case 'fade':
        return `will-change:opacity;animation:pptxFadeIn ${duration}s ease-out both;`;
      case 'cut':
        return `opacity:0;will-change:opacity;animation:pptxFadeIn ${duration}s steps(1, end) both;`;
      case 'push':
        return buildSlideTranslateTransitionStyle(duration, getTransitionTranslateOffsets(slideTransitionInfo.direction, 'right'));
      case 'wipe':
        return buildSlideClipTransitionStyle(duration, getDirectionalClipAnimationName(slideTransitionInfo.direction, 'left'));
      case 'split':
        return buildSlideClipTransitionStyle(duration, getSplitClipAnimationName(slideTransitionInfo.orientation));
      case 'reveal':
        return buildSlideClipTransitionStyle(duration, getDirectionalClipAnimationName(slideTransitionInfo.direction, 'top'));
      case 'random':
        return 'will-change:opacity;';
      case 'shape':
        return buildSlideClipTransitionStyle(duration, getShapeClipAnimationName(slideTransitionInfo.pattern));
      case 'uncover':
        return buildSlideClipTransitionStyle(duration, getDirectionalClipAnimationName(slideTransitionInfo.direction, 'left'));
      case 'cover':
        return buildSlideTranslateTransitionStyle(duration, getTransitionTranslateOffsets(slideTransitionInfo.direction, 'right'));
      default:
        return '';
    }
  },

  extractSlideTransitionShapeRenderInfo(previousSlidePart, slidePart, slideTransitionInfo) {
    const empty = { animatedShapeStylesByShapeId: new Map(), pictureOverlaysByShapeId: new Map() };
    if (previousSlidePart == null || !eqIgnoreCase(slideTransitionInfo.type, 'morph')) return empty;

    const previousShapeInfos = this._buildMorphShapeInfos(previousSlidePart);
    const currentShapeInfos = this._buildMorphShapeInfos(slidePart);
    if (currentShapeInfos.length === 0) return empty;

    const previousShapeInfosById = new Map();
    const previousShapeInfosByMatchKey = new Map();
    for (const info of previousShapeInfos) {
      if (!previousShapeInfosById.has(info.shapeId)) previousShapeInfosById.set(info.shapeId, []);
      previousShapeInfosById.get(info.shapeId).push(info);
      if (!previousShapeInfosByMatchKey.has(info.matchKey)) previousShapeInfosByMatchKey.set(info.matchKey, []);
      previousShapeInfosByMatchKey.get(info.matchKey).push(info);
    }

    const stylesByShapeId = new Map();
    const pictureOverlaysByShapeId = new Map();
    const matchedCurrentShapeIds = new Set();
    const matchedPreviousShapeIds = new Set();

    for (const currentShapeInfo of currentShapeInfos) {
      const previousMatch = this._findPreviousMorphMatch(
        currentShapeInfo,
        previousShapeInfosById,
        previousShapeInfosByMatchKey,
        matchedPreviousShapeIds
      );
      if (previousMatch == null) continue;

      matchedCurrentShapeIds.add(currentShapeInfo.shapeId);
      matchedPreviousShapeIds.add(previousMatch.shapeId);

      const overlay =
        this._buildMorphPictureOverlayInfo(currentShapeInfo, previousMatch, slideTransitionInfo.durationMs) ||
        this._buildMorphShapeOverlayInfo(currentShapeInfo, previousMatch, previousSlidePart, slideTransitionInfo.durationMs);
      if (overlay != null) pictureOverlaysByShapeId.set(currentShapeInfo.shapeId, overlay);

      const translateX = centerX(previousMatch) - centerX(currentShapeInfo);
      const translateY = centerY(previousMatch) - centerY(currentShapeInfo);
      const rotationOffsetDeg = this._normalizeMorphRotationOffset(previousMatch.rotationDeg - currentShapeInfo.rotationDeg);
      const scaleX =
        this._getMorphScaleFactor(previousMatch.widthPx, currentShapeInfo.widthPx) *
        this._getMorphTransformScaleFactor(previousMatch.transformScaleX, currentShapeInfo.transformScaleX);
      const scaleY =
        this._getMorphScaleFactor(previousMatch.heightPx, currentShapeInfo.heightPx) *
        this._getMorphTransformScaleFactor(previousMatch.transformScaleY, currentShapeInfo.transformScaleY);
      if (
        Math.abs(translateX) < 0.01 &&
        Math.abs(translateY) < 0.01 &&
        Math.abs(rotationOffsetDeg) < 0.01 &&
        Math.abs(scaleX - 1.0) < 0.001 &&
        Math.abs(scaleY - 1.0) < 0.001
      )
        continue;

      stylesByShapeId.set(
        currentShapeInfo.shapeId,
        this._buildMorphTransformStyle(translateX, translateY, rotationOffsetDeg, scaleX, scaleY, slideTransitionInfo.durationMs)
      );
    }

    for (const currentShapeInfo of currentShapeInfos) {
      if (matchedCurrentShapeIds.has(currentShapeInfo.shapeId)) continue;
      stylesByShapeId.set(currentShapeInfo.shapeId, this._buildMorphFadeInStyle(slideTransitionInfo.durationMs));
    }

    return {
      animatedShapeStylesByShapeId: stylesByShapeId,
      pictureOverlaysByShapeId,
    };
  },

  _findPreviousMorphMatch(currentShapeInfo, previousShapeInfosById, previousShapeInfosByMatchKey, matchedPreviousShapeIds) {
    if (previousShapeInfosById.has(currentShapeInfo.shapeId)) {
      const candidates = previousShapeInfosById
        .get(currentShapeInfo.shapeId)
        .filter((m) => !matchedPreviousShapeIds.has(m.shapeId) && this._areMorphMatchesCompatible(currentShapeInfo, m));
      if (candidates.length === 1) return candidates[0];
    }

    if (!previousShapeInfosByMatchKey.has(currentShapeInfo.matchKey)) return null;

    const unmatched = previousShapeInfosByMatchKey
      .get(currentShapeInfo.matchKey)
      .filter((m) => !matchedPreviousShapeIds.has(m.shapeId));
    return unmatched.length === 1 ? unmatched[0] : null;
  },

  _areMorphMatchesCompatible(currentShapeInfo, previousShapeInfo) {
    return (
      !!currentShapeInfo.stableMatchKey &&
      currentShapeInfo.stableMatchKey.trim() !== '' &&
      currentShapeInfo.stableMatchKey === previousShapeInfo.stableMatchKey
    );
  },

  buildSlideTransitionJson(slideTransitionInfo) {
    if (!slideTransitionInfo.type || slideTransitionInfo.type.trim() === '') return null;
    return { type: slideTransitionInfo.type, durationMs: slideTransitionInfo.durationMs };
  },

  _getSupportedSlideTransitionInfo(transitionElement) {
    for (const childElement of childElements(transitionElement)) {
      const localName = childElement.localName;
      if (eqIgnoreCase(localName, 'morph')) return { type: 'morph' };
      if (eqIgnoreCase(localName, 'fade')) return { type: 'fade' };
      if (eqIgnoreCase(localName, 'cut')) return { type: 'cut' };
      if (eqIgnoreCase(localName, 'push'))
        return { type: 'push', direction: normalizeDirectionalTransitionValue(attr(childElement, 'dir')) };
      if (eqIgnoreCase(localName, 'wipe'))
        return { type: 'wipe', direction: normalizeDirectionalTransitionValue(attr(childElement, 'dir')) };
      if (eqIgnoreCase(localName, 'split'))
        return {
          type: 'split',
          direction: normalizeSplitMode(attr(childElement, 'dir')),
          orientation: normalizeTransitionOrientation(attr(childElement, 'orient')),
        };
      if (eqIgnoreCase(localName, 'reveal'))
        return { type: 'reveal', direction: normalizeDirectionalTransitionValue(attr(childElement, 'dir')) };
      if (eqIgnoreCase(localName, 'randomBar'))
        return { type: 'random', orientation: normalizeTransitionOrientation(attr(childElement, 'dir')) };
      if (eqIgnoreCase(localName, 'circle')) return { type: 'shape', pattern: 'circle' };
      if (eqIgnoreCase(localName, 'pull'))
        return { type: 'uncover', direction: normalizeDirectionalTransitionValue(attr(childElement, 'dir')) };
      if (eqIgnoreCase(localName, 'cover'))
        return { type: 'cover', direction: normalizeDirectionalTransitionValue(attr(childElement, 'dir')) };
    }
    return null;
  },

  _getExplicitTransitionDurationMs(transitionElement) {
    const durationValue = attr(transitionElement, 'dur', (ns) => ns === NS.P14);
    const n = parseInt(durationValue, 10);
    return Number.isNaN(n) ? 0 : n;
  },

  _buildMorphShapeInfos(slidePart) {
    const shapeTree = shapeTreeOf(slidePart.root);
    if (shapeTree == null) return [];

    const shapeInfos = [];
    for (const shape of elements(shapeTree, 'sp', NS.P)) {
      const spPr = shapePropsOf(shape);
      if (spPr == null) continue;

      const ph = firstChild(firstChild(firstChild(shape, 'nvSpPr'), 'nvPr'), 'ph');

      const shapeId = this.getPptxShapeId(shape) || '';
      if (shapeId.trim() === '') continue;

      let [left, top, width, height] = this.getTransform(spPr);
      if (width <= 0 || height <= 0) {
        if (ph != null) {
          const resolved = this.resolvePhTransform(ph, slidePart);
          if (resolved != null) [left, top, width, height] = resolved;
        }
      }

      if (width <= 0 || height <= 0) continue;

      const geometryType = this._getMorphShapeGeometryType(spPr);
      if (geometryType.trim() === '') continue;

      const [transformScaleX, transformScaleY] = this._getMorphTransformScale(firstChild(spPr, 'xfrm'));

      const stableMatchKey = buildMorphStableShapeMatchKey(geometryType, getNormalizedMorphShapeText(shape));

      shapeInfos.push({
        shapeId,
        stableMatchKey,
        matchKey: buildMorphMatchKey(geometryType, width, height, getNormalizedMorphShapeText(shape)),
        isPicture: false,
        appearanceKey: getMorphShapeAppearanceKey(shape),
        renderAssetPath: '',
        leftPx: left,
        topPx: top,
        widthPx: width,
        heightPx: height,
        rotationDeg: this._getShapeRotationDegrees(firstChild(spPr, 'xfrm')),
        transformScaleX,
        transformScaleY,
      });
    }

    for (const picture of elements(shapeTree, 'pic', NS.P)) {
      const spPr = shapePropsOf(picture);
      if (spPr == null) continue;

      const shapeId = this.getPptxShapeId(picture) || '';
      if (shapeId.trim() === '') continue;

      let [left, top, width, height] = this.getTransform(spPr);
      if (width <= 0 || height <= 0) continue;

      const [transformScaleX, transformScaleY] = this._getMorphTransformScale(firstChild(spPr, 'xfrm'));

      const imagePart = this.resolveImagePart(slidePart, firstChild(firstChild(picture, 'blipFill'), 'blip'));
      const cNvPr = firstChild(firstChild(picture, 'nvPicPr'), 'cNvPr');
      const stableMatchKey = this._buildMorphPictureMatchKey(
        imagePart,
        this.getStringValue(attr(cNvPr, 'name')),
        this.getStringValue(attr(cNvPr, 'descr'))
      );
      if (stableMatchKey.trim() === '') continue;

      shapeInfos.push({
        shapeId,
        stableMatchKey,
        matchKey: stableMatchKey,
        isPicture: true,
        appearanceKey: this._getMorphImageAppearanceKey(imagePart),
        renderAssetPath: imagePart != null ? this.saveImageFile(imagePart, false) : '',
        leftPx: left,
        topPx: top,
        widthPx: width,
        heightPx: height,
        rotationDeg: this._getShapeRotationDegrees(firstChild(spPr, 'xfrm')),
        transformScaleX,
        transformScaleY,
      });
    }

    return shapeInfos;
  },

  _getMorphShapeGeometryType(spPr) {
    const prstGeom = firstChild(spPr, 'prstGeom');
    const geometryType = prstGeom ? attr(prstGeom, 'prst') || '' : '';
    if (geometryType.trim() !== '') return geometryType;
    return firstChild(spPr, 'custGeom') == null ? 'rect' : '';
  },

  _buildMorphPictureMatchKey(imagePart, pictureName, pictureDescription) {
    const imageContentKey = this._getMorphImageContentKey(imagePart);
    const normalizedPictureName = pictureName.trim().toLowerCase();
    const normalizedPictureDescription = pictureDescription.trim().toLowerCase();
    if (imageContentKey.length > 0) {
      return ['picture', imageContentKey].join('|');
    }
    if (normalizedPictureName.length === 0 && normalizedPictureDescription.length === 0) return '';
    return ['picture', normalizedPictureName, normalizedPictureDescription].join('|');
  },

  _getMorphImageContentKey(imagePart) {
    if (imagePart == null) return '';
    try {
      const imageBytes = Buffer.from(imagePart.getBytes());
      if (eqIgnoreCase(imagePart.contentType, 'image/svg+xml')) {
        const svgContentKey = getMorphSvgContentKey(imageBytes);
        if (svgContentKey.length > 0) return svgContentKey;
      }
      return sha256Hex(imageBytes);
    } catch {
      return '';
    }
  },

  _getMorphImageAppearanceKey(imagePart) {
    if (imagePart == null) return '';
    try {
      return sha256Hex(Buffer.from(imagePart.getBytes()));
    } catch {
      return '';
    }
  },

  _getShapeRotationDegrees(transform) {
    const rot = attr(transform, 'rot');
    return (rot != null ? Number(rot) : 0) / 60000.0;
  },

  _getMorphTransformScale(transform) {
    const flipH = attr(transform, 'flipH') === '1' || (attr(transform, 'flipH') || '').toLowerCase() === 'true';
    const flipV = attr(transform, 'flipV') === '1' || (attr(transform, 'flipV') || '').toLowerCase() === 'true';

    let scaleX = 1.0;
    let scaleY = 1.0;
    if (flipV && !flipH) {
      scaleX = -1.0;
      scaleY = -1.0;
    } else {
      if (flipH) scaleX = -1.0;
      if (flipV) scaleY = -1.0;
    }
    return [scaleX, scaleY];
  },

  _normalizeMorphRotationOffset(rotationOffsetDeg) {
    let normalized = rotationOffsetDeg % 360.0;
    if (normalized > 180.0) normalized -= 360.0;
    else if (normalized < -180.0) normalized += 360.0;
    return normalized;
  },

  _getMorphScaleFactor(previousSizePx, currentSizePx) {
    return currentSizePx > 0.01 ? previousSizePx / currentSizePx : 1.0;
  },

  _getMorphTransformScaleFactor(previousScale, currentScale) {
    return Math.abs(previousScale) > 0.001 && Math.abs(currentScale) > 0.001 ? previousScale / currentScale : 1.0;
  },

  _buildMorphPictureOverlayInfo(currentShapeInfo, previousShapeInfo, durationMs) {
    if (!currentShapeInfo.isPicture || !previousShapeInfo.isPicture) return null;
    if (!previousShapeInfo.renderAssetPath || previousShapeInfo.renderAssetPath.trim() === '') return null;
    if (
      !currentShapeInfo.appearanceKey ||
      currentShapeInfo.appearanceKey.trim() === '' ||
      currentShapeInfo.appearanceKey === previousShapeInfo.appearanceKey
    ) {
      return null;
    }
    return {
      previousImagePath: previousShapeInfo.renderAssetPath,
      previousImageStyle: buildMorphOverlayFadeStyle(this, durationMs),
      previousSvgMarkup: '',
    };
  },

  _buildMorphShapeOverlayInfo(currentShapeInfo, previousShapeInfo, previousSlidePart, durationMs) {
    if (currentShapeInfo.isPicture || previousShapeInfo.isPicture) return null;
    if (
      !currentShapeInfo.appearanceKey ||
      currentShapeInfo.appearanceKey.trim() === '' ||
      currentShapeInfo.appearanceKey === previousShapeInfo.appearanceKey
    ) {
      return null;
    }

    const previousShape = this._findMorphShape(previousSlidePart, previousShapeInfo.shapeId);
    if (previousShape == null) return null;

    const overlayMarkup = this._buildMorphShapeOverlayMarkup(previousShape, previousSlidePart, previousShapeInfo, currentShapeInfo, durationMs);
    if (!overlayMarkup || overlayMarkup.trim() === '') return null;

    return { previousImagePath: '', previousImageStyle: '', previousSvgMarkup: overlayMarkup };
  },

  _findMorphShape(slidePart, shapeId) {
    const shapeTree = shapeTreeOf(slidePart.root);
    if (shapeTree == null) return null;
    return elements(shapeTree, 'sp', NS.P).find((shape) => this.getPptxShapeId(shape) === shapeId) || null;
  },

  _buildMorphShapeOverlayMarkup(previousShape, previousSlidePart, previousShapeInfo, currentShapeInfo, durationMs) {
    if (
      previousShapeInfo.widthPx <= 0.01 ||
      previousShapeInfo.heightPx <= 0.01 ||
      currentShapeInfo.widthPx <= 0.01 ||
      currentShapeInfo.heightPx <= 0.01
    ) {
      return '';
    }

    const defsSb = new StringBuilder();
    const bodySb = new StringBuilder();
    this.renderShape(bodySb, defsSb, previousShape, previousSlidePart);
    if (bodySb.length === 0) return '';

    const currentWidth = fmtTrim(currentShapeInfo.widthPx, 3);
    const currentHeight = fmtTrim(currentShapeInfo.heightPx, 3);
    const viewBox = [
      fmtTrim(previousShapeInfo.leftPx, 3),
      fmtTrim(previousShapeInfo.topPx, 3),
      fmtTrim(previousShapeInfo.widthPx, 3),
      fmtTrim(previousShapeInfo.heightPx, 3),
    ].join(' ');
    const overlayStyleAttr = this.buildSvgStyleAttribute(buildMorphOverlayFadeStyle(this, durationMs));

    const sb = new StringBuilder();
    sb.appendLine(
      `<svg x="0" y="0" width="${currentWidth}" height="${currentHeight}" viewBox="${viewBox}" preserveAspectRatio="none" overflow="visible"${overlayStyleAttr}>`
    );
    if (defsSb.length > 0) {
      sb.appendLine('<defs>');
      sb.append(defsSb.toString());
      sb.appendLine('</defs>');
    }
    sb.append(bodySb.toString());
    sb.appendLine('</svg>');
    return sb.toString();
  },

  _buildMorphTransformStyle(translateX, translateY, rotationOffsetDeg, scaleX, scaleY, durationMs) {
    const duration = fmtTrim(this.toCssDurationSeconds(durationMs, 0.5), 3);
    const translateXValue = fmtTrim(translateX, 2);
    const translateYValue = fmtTrim(translateY, 2);
    const rotationOffset = fmtTrim(rotationOffsetDeg, 2);
    const scaleXValue = fmtTrim(scaleX, 4);
    const scaleYValue = fmtTrim(scaleY, 4);
    return `will-change:transform;transform-box:fill-box;transform-origin:center center;animation:pptxMorphTransform ${duration}s ease-in-out both;--pptx-morph-translate-x:${translateXValue}px;--pptx-morph-translate-y:${translateYValue}px;--pptx-morph-rotate-from:${rotationOffset}deg;--pptx-morph-scale-x:${scaleXValue};--pptx-morph-scale-y:${scaleYValue};`;
  },

  _buildMorphFadeInStyle(durationMs) {
    const duration = fmtTrim(this.toCssDurationSeconds(durationMs, 0.5), 3);
    return `will-change:opacity;animation:pptxFadeIn ${duration}s ease-out both;`;
  },
};

// ── module-scope helpers (no converter state) ──────────────────────────────

function buildSlideTranslateTransitionStyle(duration, offsets) {
  return `will-change:transform;animation:pptxSlideTranslateIn ${duration}s ease-in-out both;--pptx-slide-translate-from-x:${offsets[0]};--pptx-slide-translate-from-y:${offsets[1]};`;
}

function buildSlideClipTransitionStyle(duration, animationName, timingFunction = 'ease-in-out') {
  return `will-change:opacity,clip-path;animation:${animationName} ${duration}s ${timingFunction} both;`;
}

function buildMorphOverlayFadeStyle(ctx, durationMs) {
  const duration = fmtTrim(ctx.toCssDurationSeconds(durationMs, 0.5), 3);
  return `will-change:opacity;animation:pptxMorphFadeOut ${duration}s ease-in-out both;`;
}

function centerX(info) {
  return info.leftPx + info.widthPx / 2.0;
}
function centerY(info) {
  return info.topPx + info.heightPx / 2.0;
}

function getNormalizedMorphShapeText(shape) {
  const parts = descendants(textBodyOf(shape), 't', NS.A)
    .map((t) => (innerText(t) || '').trim())
    .filter((t) => t.length > 0);
  return parts.join(' ').trim().toLowerCase();
}

function buildMorphStableShapeMatchKey(geometryType, text) {
  return [geometryType.trim().toLowerCase(), text].join('|');
}

function buildMorphMatchKey(geometryType, width, height, text) {
  return [geometryType.trim().toLowerCase(), String(bankRound(width)), String(bankRound(height)), text].join('|');
}

function getMorphShapeAppearanceKey(shape) {
  const spPr = shapePropsOf(shape);
  const shapeProperties = spPr
    ? childElements(spPr)
        .filter((child) => child.localName !== 'xfrm')
        .map((child) => outerXml(child))
        .join('')
    : '';
  const styleEl = firstChild(shape, 'style');
  const shapeStyle = styleEl ? outerXml(styleEl) : '';
  const txBody = textBodyOf(shape);
  const textBody = txBody ? outerXml(txBody) : '';
  if (shapeProperties.length === 0 && shapeStyle.length === 0 && textBody.length === 0) return '';

  return sha256Hex(Buffer.from([shapeProperties, shapeStyle, textBody].join('|'), 'utf8'));
}

function getMorphSvgContentKey(svgBytes) {
  try {
    let text = Buffer.from(svgBytes).toString('utf8');
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    const doc = _svgParser.parseFromString(text, 'text/xml');
    const root = doc.documentElement;
    const rootId = ((root ? attr(root, 'id') : null) || '').trim().toLowerCase();
    const viewBox = ((root ? attr(root, 'viewBox') : null) || '').trim();
    const pathData = descendants(root)
      .filter((el) => (el.localName || '').toLowerCase() === 'path')
      .map((el) => (attr(el, 'd') || '').trim())
      .filter((d) => d.length > 0)
      .join('|');
    if (rootId.length === 0 && viewBox.length === 0 && pathData.length === 0) return '';
    return ['svg', rootId, viewBox, pathData].join('|');
  } catch {
    return '';
  }
}
