'use strict';

// Port of PptxToHtml/pptx/AnimationStageHelper.cs

const { childElements, descendants, attr, innerText, NS } = require('../ooxml/dom');
const { fmtTrim } = require('../util/format');

const PresentationNamespace = NS.P;

function isP(el, localName) {
  return el != null && el.namespaceURI === PresentationNamespace && el.localName === localName;
}

// FindChildElement(parent, localName): first direct child that is p:localName.
function pFirstChild(parent, localName) {
  if (parent == null) return null;
  for (const child of childElements(parent)) {
    if (isP(child, localName)) return child;
  }
  return null;
}

// Descendants filtered to p:localName.
function pDescendants(el, localName) {
  return descendants(el, localName, PresentationNamespace);
}

function getAttrValue(el, localName) {
  return attr(el, localName);
}

function eqIgnoreCase(a, b) {
  return (a || '').toLowerCase() === (b || '').toLowerCase();
}

// Empty singletons.
const EmptyAnimatedShapeStyles = new Map();

class RawShapeAnimationInfo {
  constructor() {
    this.entranceFilter = '';
    this.entranceTransition = '';
    this.entranceDurationMs = 0;
    this.numericAnimationsByAttributeName = new Map(); // key lowercased for OIC
  }
}

class RawSlideAnimationStage {
  constructor() {
    this.targetShapeIds = new Set();
    this.showCandidateShapeIds = new Set();
    this.hideCandidateShapeIds = new Set();
    this.shapeAnimationsByShapeId = new Map();
  }
}

module.exports = {
  extractSlideAnimationInfo(slidePart) {
    const slide = slidePart.root;
    if (slide == null) return { initiallyHiddenShapeIds: [], stages: [] };

    const timing = childElements(slide).find((el) => isP(el, 'timing'));
    if (timing == null) return { initiallyHiddenShapeIds: [], stages: [] };

    const mainSequence = descendants(timing).find(
      (el) => isP(el, 'seq') && eqIgnoreCase(getAttrValue(pFirstChild(el, 'cTn'), 'nodeType'), 'mainSeq')
    );
    if (mainSequence == null) return { initiallyHiddenShapeIds: [], stages: [] };

    const childTimeNodeList = pFirstChild(pFirstChild(mainSequence, 'cTn'), 'childTnLst');
    if (childTimeNodeList == null) return { initiallyHiddenShapeIds: [], stages: [] };

    const rawStages = [];
    for (const stageElement of childElements(childTimeNodeList).filter((el) => isP(el, 'par'))) {
      const rawStage = this._parseRawSlideAnimationStage(stageElement);
      if (rawStage.targetShapeIds.size > 0 && this._stageHasRenderedEffect(stageElement)) rawStages.push(rawStage);
    }

    if (rawStages.length === 0) return { initiallyHiddenShapeIds: [], stages: [] };

    return this._buildSlideAnimationInfo(rawStages);
  },

  _buildSlideAnimationInfo(rawStages) {
    const firstVisibilityActions = new Map(); // shapeId -> 'None'|'Show'|'Hide'
    for (const rawStage of rawStages) {
      for (const shapeId of rawStage.targetShapeIds) {
        if (firstVisibilityActions.has(shapeId)) continue;
        if (rawStage.hideCandidateShapeIds.has(shapeId)) {
          firstVisibilityActions.set(shapeId, 'Hide');
          continue;
        }
        if (rawStage.showCandidateShapeIds.has(shapeId)) {
          firstVisibilityActions.set(shapeId, 'Show');
          continue;
        }
        firstVisibilityActions.set(shapeId, 'None');
      }
    }

    const initialHiddenShapeIds = new Set();
    for (const [key, value] of firstVisibilityActions) {
      if (value === 'Show') initialHiddenShapeIds.add(key);
    }

    const hiddenShapeIds = new Set(initialHiddenShapeIds);
    const activePersistentAnimationPartsByShapeId = new Map();
    const stages = [];

    for (let index = 0; index < rawStages.length; index++) {
      const rawStage = rawStages[index];
      const rawShapeAnimationCssInfos = this._buildAnimatedShapeCssInfos(rawStage);

      for (const shapeId of rawStage.targetShapeIds) {
        if (rawStage.hideCandidateShapeIds.has(shapeId)) {
          hiddenShapeIds.add(shapeId);
          activePersistentAnimationPartsByShapeId.delete(shapeId);
          continue;
        }
        if (rawStage.showCandidateShapeIds.has(shapeId)) hiddenShapeIds.delete(shapeId);
      }

      const stageAnimationPartsByShapeId = new Map();
      for (const [k, v] of activePersistentAnimationPartsByShapeId) {
        stageAnimationPartsByShapeId.set(k, v.slice());
      }

      for (const [shapeId, cssInfo] of rawShapeAnimationCssInfos) {
        if (cssInfo.currentStageAnimationParts.length > 0) {
          stageAnimationPartsByShapeId.set(
            shapeId,
            this._mergeAnimationParts(
              cssInfo.currentStageAnimationParts,
              stageAnimationPartsByShapeId.has(shapeId) ? stageAnimationPartsByShapeId.get(shapeId) : null
            )
          );
        }

        if (cssInfo.persistentAnimationParts.length > 0) {
          activePersistentAnimationPartsByShapeId.set(shapeId, cssInfo.persistentAnimationParts.slice());
        }
      }

      stages.push({
        stageNumber: index + 1,
        hiddenShapeIdsAfterStage: this._sortShapeIds(hiddenShapeIds),
        animatedShapeStylesByShapeId: this._buildAnimatedShapeStyles(stageAnimationPartsByShapeId),
      });
    }

    return {
      initiallyHiddenShapeIds: this._sortShapeIds(initialHiddenShapeIds),
      stages,
    };
  },

  _parseRawSlideAnimationStage(stageElement) {
    const rawStage = new RawSlideAnimationStage();
    const commonTimeNode = pFirstChild(stageElement, 'cTn');
    if (commonTimeNode == null || !this._hasIndefiniteStartCondition(commonTimeNode)) return rawStage;

    for (const shapeId of this._getTargetShapeIds(stageElement)) rawStage.targetShapeIds.add(shapeId);

    for (const timingNode of pDescendants(stageElement, 'cTn')) {
      const presetClass = getAttrValue(timingNode, 'presetClass');
      if (eqIgnoreCase(presetClass, 'entr')) {
        for (const id of this._getTargetShapeIds(timingNode)) rawStage.showCandidateShapeIds.add(id);
      } else if (eqIgnoreCase(presetClass, 'exit')) {
        for (const id of this._getTargetShapeIds(timingNode)) rawStage.hideCandidateShapeIds.add(id);
      }
    }

    for (const setElement of pDescendants(stageElement, 'set')) {
      const targetsVisibility = pDescendants(setElement, 'attrName')
        .map((el) => innerText(el))
        .some((name) => eqIgnoreCase(name, 'style.visibility'));
      if (!targetsVisibility) continue;

      const strVal = pDescendants(setElement, 'strVal')[0];
      const value = getAttrValue(strVal, 'val');
      if (eqIgnoreCase(value, 'visible')) {
        for (const id of this._getTargetShapeIds(setElement)) rawStage.showCandidateShapeIds.add(id);
      } else if (eqIgnoreCase(value, 'hidden')) {
        for (const id of this._getTargetShapeIds(setElement)) rawStage.hideCandidateShapeIds.add(id);
      }
    }

    for (const animationEffect of pDescendants(stageElement, 'animEffect')) {
      const transition = getAttrValue(animationEffect, 'transition');
      const filter = getAttrValue(animationEffect, 'filter') || '';
      const durationMs = this._getDurationMs(pFirstChild(pFirstChild(animationEffect, 'cBhvr'), 'cTn'));
      if (eqIgnoreCase(transition, 'in')) {
        for (const id of this._getTargetShapeIds(animationEffect)) rawStage.showCandidateShapeIds.add(id);
      } else if (eqIgnoreCase(transition, 'out')) {
        for (const id of this._getTargetShapeIds(animationEffect)) rawStage.hideCandidateShapeIds.add(id);
      }

      for (const shapeId of this._getTargetShapeIds(animationEffect)) {
        if (!eqIgnoreCase(transition, 'in')) continue;
        const shapeAnimation = this._getOrCreateShapeAnimationInfo(rawStage, shapeId);
        shapeAnimation.entranceTransition = transition || '';
        if (filter && filter.trim() !== '') shapeAnimation.entranceFilter = filter;
        if (durationMs > 0) shapeAnimation.entranceDurationMs = durationMs;
      }
    }

    for (const animation of pDescendants(stageElement, 'anim')) {
      const attributeName =
        pDescendants(animation, 'attrName')
          .map((el) => innerText(el))
          .find((name) => name && name.trim() !== '') || '';
      if (attributeName.trim() === '') continue;

      const expressions = pDescendants(animation, 'tav')
        .map((tav) => this._getTimingValueExpression(tav))
        .filter((expr) => expr && expr.trim() !== '');
      if (expressions.length === 0) continue;

      const hasSineExpression = expressions.some(
        (e) => e.toLowerCase().includes('sin(') || e.toLowerCase().includes('pi')
      );
      const hasTimeFormula = expressions.some((e) => e.includes('$'));
      const distinctCount = new Set(expressions.map((e) => e.toLowerCase())).size;
      const changesValue = hasTimeFormula || distinctCount > 1;
      const durationMs = this._getDurationMs(pFirstChild(pFirstChild(animation, 'cBhvr'), 'cTn'));
      const repeatCount = this._getNearestRepeatCount(animation);

      for (const shapeId of this._getTargetShapeIds(animation)) {
        const shapeAnimation = this._getOrCreateShapeAnimationInfo(rawStage, shapeId);
        shapeAnimation.numericAnimationsByAttributeName.set(attributeName.toLowerCase(), {
          durationMs,
          repeatCount,
          expressions,
          changesValue,
          hasSineExpression,
        });
      }
    }

    return rawStage;
  },

  _buildAnimatedShapeCssInfos(rawStage) {
    const animationCssInfosByShapeId = new Map();
    for (const [shapeId, shapeAnimationInfo] of rawStage.shapeAnimationsByShapeId) {
      const animationCssInfo = this._buildShapeAnimationCssInfo(shapeAnimationInfo);
      if (
        animationCssInfo.currentStageAnimationParts.length === 0 &&
        animationCssInfo.persistentAnimationParts.length === 0
      ) {
        continue;
      }
      animationCssInfosByShapeId.set(shapeId, animationCssInfo);
    }
    return animationCssInfosByShapeId;
  },

  _buildAnimatedShapeStyles(animationPartsByShapeId) {
    if (animationPartsByShapeId.size === 0) return EmptyAnimatedShapeStyles;

    const result = new Map();
    for (const [key, parts] of animationPartsByShapeId) {
      const style = this._buildShapeAnimationCss(parts);
      if (style && style.trim() !== '') result.set(key, style);
    }
    return result;
  },

  _buildShapeAnimationCssInfo(shapeAnimationInfo) {
    const currentStageAnimationParts = [];
    const persistentAnimationParts = [];

    const entranceDurationSeconds = this.toCssDurationSeconds(shapeAnimationInfo.entranceDurationMs, 0.5);
    const entranceDuration = fmtTrim(entranceDurationSeconds, 3);
    if (shapeAnimationInfo.entranceFilter.toLowerCase().includes('fade')) {
      currentStageAnimationParts.push(`pptxFadeIn ${entranceDuration}s ease-out both`);
    } else if (shapeAnimationInfo.entranceFilter.toLowerCase().includes('randombar(horizontal)')) {
      currentStageAnimationParts.push(`pptxHorizontalBarsIn ${entranceDuration}s steps(8, end) both`);
    } else if (eqIgnoreCase(shapeAnimationInfo.entranceTransition, 'in')) {
      currentStageAnimationParts.push(`pptxFadeIn ${entranceDuration}s ease-out both`);
    }

    const widthAnimation = shapeAnimationInfo.numericAnimationsByAttributeName.get('ppt_w');
    if (widthAnimation && widthAnimation.changesValue) {
      const widthDurationSeconds = this.toCssDurationSeconds(
        widthAnimation.durationMs,
        entranceDurationSeconds > 0 ? entranceDurationSeconds : 1.0
      );
      const widthDuration = fmtTrim(widthDurationSeconds, 3);
      const isInfiniteLoop = this._isIndefiniteRepeat(widthAnimation.repeatCount);
      const repeatValue = isInfiniteLoop ? 'infinite' : '1';
      const widthAnimationCss = this._isSwivelWidthAnimation(widthAnimation)
        ? `${isInfiniteLoop ? 'pptxRotateYSwivelLoop' : 'pptxRotateYSwivel'} ${widthDuration}s linear ${repeatValue} both`
        : widthAnimation.hasSineExpression
        ? `${isInfiniteLoop ? 'pptxScaleXSineLoop' : 'pptxScaleXSine'} ${widthDuration}s linear ${repeatValue} both`
        : `${isInfiniteLoop ? 'pptxScaleXGrowLoop' : 'pptxScaleXGrow'} ${widthDuration}s ease-in-out ${repeatValue} both`;
      currentStageAnimationParts.push(widthAnimationCss);

      if (isInfiniteLoop) persistentAnimationParts.push(widthAnimationCss);
    }

    return { currentStageAnimationParts, persistentAnimationParts };
  },

  _buildShapeAnimationCss(animationParts) {
    if (animationParts.length === 0) return '';
    return (
      'will-change:opacity,transform,clip-path;transform-box:fill-box;transform-origin:center center;transform-style:preserve-3d;backface-visibility:visible;animation:' +
      animationParts.join(',') +
      ';'
    );
  },

  _isSwivelWidthAnimation(widthAnimation) {
    return widthAnimation.expressions.some((e) => e.toLowerCase().includes('#ppt_w*sin(2.5*pi*$)'));
  },

  _mergeAnimationParts(primaryAnimationParts, secondaryAnimationParts) {
    const merged = [];
    for (const part of primaryAnimationParts.concat(secondaryAnimationParts || [])) {
      if (merged.includes(part)) continue;
      merged.push(part);
    }
    return merged;
  },

  // A media-playback node (audio/video "playFrom" command, presetClass="mediacall") is stored
  // in the timing tree as a click stage, but it produces no visual change. Only treat a par as a
  // real stage when it carries an effect this converter actually renders: an entrance/exit
  // (<p:set> visibility toggle / <p:animEffect>) or a numeric emphasis (<p:anim>). Without one,
  // the emitted sub-slide would be a byte-for-byte duplicate of the main slide (e.g. audio-2-files
  // was producing slideN.html plus an identical slideN_1.html per audio clip).
  _stageHasRenderedEffect(stageElement) {
    return descendants(stageElement).some(
      (element) => isP(element, 'set') || isP(element, 'animEffect') || isP(element, 'anim')
    );
  },

  _hasIndefiniteStartCondition(commonTimeNode) {
    const startConditionList = pFirstChild(commonTimeNode, 'stCondLst');
    if (startConditionList == null) return false;
    return childElements(startConditionList)
      .filter((el) => isP(el, 'cond'))
      .map((el) => getAttrValue(el, 'delay'))
      .some((delay) => eqIgnoreCase(delay, 'indefinite'));
  },

  _getTargetShapeIds(element) {
    const seen = new Set();
    const out = [];
    for (const child of descendants(element, 'spTgt', PresentationNamespace)) {
      const shapeId = getAttrValue(child, 'spid');
      if (shapeId && shapeId.trim() !== '' && !seen.has(shapeId)) {
        seen.add(shapeId);
        out.push(shapeId);
      }
    }
    return out;
  },

  _getOrCreateShapeAnimationInfo(rawStage, shapeId) {
    let info = rawStage.shapeAnimationsByShapeId.get(shapeId);
    if (!info) {
      info = new RawShapeAnimationInfo();
      rawStage.shapeAnimationsByShapeId.set(shapeId, info);
    }
    return info;
  },

  _getDurationMs(timingNode) {
    const duration = getAttrValue(timingNode, 'dur');
    const n = parseInt(duration, 10);
    return Number.isNaN(n) ? 0 : n;
  },

  _getNearestRepeatCount(element) {
    for (let current = element; current != null && current.nodeType === 1; current = current.parentNode) {
      const repeatCount = getAttrValue(current, 'repeatCount');
      if (repeatCount && repeatCount.trim() !== '') return repeatCount;
    }
    return '';
  },

  _getTimingValueExpression(timingValue) {
    const formula = getAttrValue(timingValue, 'fmla');
    if (formula && formula.trim() !== '') return formula;

    const valueElement = pFirstChild(timingValue, 'val');
    return (
      getAttrValue(pFirstChild(valueElement, 'fltVal'), 'val') ||
      getAttrValue(pFirstChild(valueElement, 'strVal'), 'val') ||
      ''
    );
  },

  _isIndefiniteRepeat(repeatCount) {
    return eqIgnoreCase(repeatCount, 'indefinite');
  },

  toCssDurationSeconds(durationMs, fallbackSeconds) {
    return durationMs > 0 ? durationMs / 1000.0 : fallbackSeconds;
  },

  _sortShapeIds(shapeIds) {
    return Array.from(shapeIds).sort((a, b) => {
      const na = /^\d+$/.test(a) ? parseInt(a, 10) : Number.MAX_SAFE_INTEGER;
      const nb = /^\d+$/.test(b) ? parseInt(b, 10) : Number.MAX_SAFE_INTEGER;
      if (na !== nb) return na - nb;
      return a < b ? -1 : a > b ? 1 : 0;
    });
  },
};
