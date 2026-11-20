'use strict';

// Port of PptxToHtml/pptx/AnimationStageHelper.cs

const { childElements, descendants, attr, attrLong, firstChild, innerText, NS } = require('../ooxml/dom');
const { placeholderOf, shapePropsOf, shapeTreeOf } = require('./elements');
const { fmtTrim, f1 } = require('../util/format');

// Animation targets are keyed by shape id ("5"), or by "<shape id>#<paragraph index>" ("3#1")
// when an effect builds a shape's text one paragraph at a time (<p:txEl><p:pRg>). The stage
// sets and style maps below carry both kinds of key; shape renderers look up shape ids and
// the text renderer looks up paragraph keys.

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
    this.motionPathsByTarget = new Map();
    this.positionAnimationsByTarget = new Map();
  }
}

// A position animation (<p:anim> on ppt_x / ppt_y, as Fly In, Float In or Crawl In use): the
// target starts where the from formulas put it and glides back to its own position. The
// formulas are fractions of the slide and may use the target's #ppt_x/#ppt_y/#ppt_w/#ppt_h.
class RawPositionAnimation {
  constructor() {
    this.fromX = null;
    this.toX = null;
    this.fromY = null;
    this.toY = null;
    this.durationMs = 0;
    this.easing = 'linear';
    this.isInfiniteLoop = false;
    // False once a ppt_x / ppt_y animation of the target in this stage can't be modelled
    // (time-based fmla, other than two keyframes, a second animation on the same axis…).
    this.isSupported = true;
  }
}

// A formula's value as constant + x*#ppt_x + y*#ppt_y + w*#ppt_w + h*#ppt_h, where the
// variables are the target's centre and size in fractions of the slide.
function linearFormula(constant, x, y, w, h) {
  return { constant, x, y, w, h };
}
function formulaIsConstant(f) {
  return f.x === 0 && f.y === 0 && f.w === 0 && f.h === 0;
}
function formulaPlus(a, b) {
  return linearFormula(a.constant + b.constant, a.x + b.x, a.y + b.y, a.w + b.w, a.h + b.h);
}
function formulaTimes(f, factor) {
  return linearFormula(f.constant * factor, f.x * factor, f.y * factor, f.w * factor, f.h * factor);
}
function formulaDividedBy(f, divisor) {
  return linearFormula(f.constant / divisor, f.x / divisor, f.y / divisor, f.w / divisor, f.h / divisor);
}

// An offset (slide px) along one axis as a function of the target's box (px):
// constant + left*left + top*top + width*width + height*height.
function offsetDependsOnBox(o) {
  return Math.abs(o.left) > 1e-9 || Math.abs(o.top) > 1e-9 || Math.abs(o.width) > 1e-9 || Math.abs(o.height) > 1e-9;
}
function offsetIsZero(o) {
  return !offsetDependsOnBox(o) && Math.abs(o.constant) < 0.05;
}
function offsetAt(o, box) {
  return o.constant + o.left * box.left + o.top * box.top + o.width * box.width + o.height * box.height;
}

const EmptyShapeBoxes = new Map();
const EmptyBox = { left: 0, top: 0, width: 0, height: 0 };

const FORMULA_TOKEN = /#ppt_[xywh]|(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]+)?|[A-Za-z]+/y;

// Reads a timing-value formula that is linear in the target's position and size, such as
// "1+#ppt_h/2" or "#ppt_y+.1": numbers, #ppt_x/#ppt_y/#ppt_w/#ppt_h, pi, e, + - * / and
// parentheses. Anything else (the time variable $, functions, a product of two variables)
// gives null.
function parseLinearFormula(text) {
  const state = { text, position: 0 };
  const value = parseFormulaSum(state);
  skipFormulaSpaces(state);
  return value != null && state.position === text.length ? value : null;
}

function parseFormulaSum(state) {
  let left = parseFormulaProduct(state);
  while (left != null) {
    skipFormulaSpaces(state);
    const c = state.text[state.position];
    if (state.position >= state.text.length || (c !== '+' && c !== '-')) break;

    const subtract = state.text[state.position++] === '-';
    const right = parseFormulaProduct(state);
    if (right == null) return null;
    left = formulaPlus(left, subtract ? formulaTimes(right, -1) : right);
  }
  return left;
}

function parseFormulaProduct(state) {
  let left = parseFormulaUnary(state);
  while (left != null) {
    skipFormulaSpaces(state);
    const c = state.text[state.position];
    if (state.position >= state.text.length || (c !== '*' && c !== '/')) break;

    const divide = state.text[state.position++] === '/';
    const right = parseFormulaUnary(state);
    if (right == null) return null;
    if (divide) left = formulaIsConstant(right) && right.constant !== 0 ? formulaDividedBy(left, right.constant) : null;
    else if (formulaIsConstant(right)) left = formulaTimes(left, right.constant);
    else left = formulaIsConstant(left) ? formulaTimes(right, left.constant) : null;
  }
  return left;
}

function parseFormulaUnary(state) {
  skipFormulaSpaces(state);
  const c = state.text[state.position];
  if (state.position < state.text.length && (c === '-' || c === '+')) {
    const negate = state.text[state.position++] === '-';
    const operand = parseFormulaUnary(state);
    return operand == null ? null : negate ? formulaTimes(operand, -1) : operand;
  }

  if (state.position < state.text.length && c === '(') {
    state.position++;
    const inner = parseFormulaSum(state);
    skipFormulaSpaces(state);
    if (inner == null || state.position >= state.text.length || state.text[state.position] !== ')') return null;
    state.position++;
    return inner;
  }

  FORMULA_TOKEN.lastIndex = state.position;
  const token = FORMULA_TOKEN.exec(state.text);
  if (token == null) return null;
  state.position += token[0].length;
  switch (token[0]) {
    case '#ppt_x':
      return linearFormula(0, 1, 0, 0, 0);
    case '#ppt_y':
      return linearFormula(0, 0, 1, 0, 0);
    case '#ppt_w':
      return linearFormula(0, 0, 0, 1, 0);
    case '#ppt_h':
      return linearFormula(0, 0, 0, 0, 1);
    case 'pi':
      return linearFormula(Math.PI, 0, 0, 0, 0);
    case 'e':
      return linearFormula(Math.E, 0, 0, 0, 0);
  }

  const first = token[0][0];
  return (first >= '0' && first <= '9') || first === '.' ? linearFormula(Number(token[0]), 0, 0, 0, 0) : null;
}

function skipFormulaSpaces(state) {
  while (state.position < state.text.length && state.text[state.position] === ' ') state.position++;
}

function formatFlyPx(valuePx) {
  return `${f1(normalizeMotionOffset(valuePx))}px`;
}

function appendBoxTerm(parts, factor, name, fallbackPx) {
  if (Math.abs(factor) <= 1e-9) return;

  parts.push(factor < 0 ? ' - ' : ' + ');
  const magnitude = Math.abs(factor);
  if (Math.abs(magnitude - 1) > 1e-9) parts.push(`${fmtTrim(magnitude, 3)} * `);
  parts.push(`var(--pptx-box-${name}, ${formatFlyPx(fallbackPx)})`);
}

function formatFlyOffset(offset, box, measured) {
  if (!measured || !offsetDependsOnBox(offset)) return formatFlyPx(offsetAt(offset, box));

  const parts = ['calc(', formatFlyPx(offset.constant)];
  appendBoxTerm(parts, offset.left, 'left', box.left);
  appendBoxTerm(parts, offset.top, 'top', box.top);
  appendBoxTerm(parts, offset.width, 'width', box.width);
  appendBoxTerm(parts, offset.height, 'height', box.height);
  parts.push(')');
  return parts.join('');
}

// How far (px) a target whose ppt_x (horizontal) or ppt_y is value sits from its own
// position, given #ppt_x = (left + width/2)/slide width, #ppt_w = width/slide width and
// likewise vertically.
function toBoxOffset(value, horizontal, slideWidthPx, slideHeightPx) {
  const axis = horizontal ? slideWidthPx : slideHeightPx;
  const x = horizontal ? value.x - 1 : value.x;
  const y = horizontal ? value.y : value.y - 1;
  const perWidth = axis / slideWidthPx;
  const perHeight = axis / slideHeightPx;
  return {
    constant: value.constant * axis,
    left: x * perWidth,
    top: y * perHeight,
    width: (x * perWidth) / 2 + value.w * perWidth,
    height: (y * perHeight) / 2 + value.h * perHeight,
  };
}

// Offset of the start point from the target's own position along one axis: zero for an axis
// that isn't animated, null when the formulas can't be read or the animation doesn't end where
// the target rests.
function getStartOffset(from, to, horizontal, slideWidthPx, slideHeightPx) {
  if (from == null || to == null) return { constant: 0, left: 0, top: 0, width: 0, height: 0 };

  const fromValue = parseLinearFormula(from);
  const toValue = parseLinearFormula(to);
  if (fromValue == null || toValue == null) return null;

  if (!offsetIsZero(toBoxOffset(toValue, horizontal, slideWidthPx, slideHeightPx))) return null;
  return toBoxOffset(fromValue, horizontal, slideWidthPx, slideHeightPx);
}

// Offsets under 0.05px print as 0.0; snapping them to zero also keeps a tiny negative value
// from printing as "-0.0" (which .NET would write, unlike f1).
function normalizeMotionOffset(offsetPx) {
  return Math.abs(offsetPx) < 0.05 ? 0 : offsetPx;
}

const MOTION_PATH_TOKEN = /[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?|[A-Za-z]/g;

// Start and end point of a motion path such as "M 0 0 L 0.25 -0.5 E" (fractions of the slide
// size), or null. PowerPoint writes M/L/C/Z/E; relative lower-case commands and H/V are read
// too, and anything else is skipped.
function parseMotionPathEnds(pathText) {
  if (pathText == null || pathText.trim() === '') return null;

  let x = 0;
  let y = 0;
  let subpathX = 0;
  let subpathY = 0;
  let start = null;
  let command = ' ';
  let args = [];
  for (const token of pathText.match(MOTION_PATH_TOKEN) || []) {
    if (/^[A-Za-z]$/.test(token)) {
      command = token;
      args = [];
      if (command === 'Z' || command === 'z') {
        x = subpathX;
        y = subpathY;
      }
      continue;
    }

    args.push(Number(token));
    const upper = command.toUpperCase();
    const arity = upper === 'M' || upper === 'L' ? 2 : upper === 'H' || upper === 'V' ? 1 : upper === 'C' ? 6 : 0;
    if (arity === 0) {
      args = [];
      continue;
    }
    if (args.length < arity) continue;

    const relative = command !== upper;
    if (upper === 'H') x = relative ? x + args[0] : args[0];
    else if (upper === 'V') y = relative ? y + args[0] : args[0];
    else {
      x = relative ? x + args[arity - 2] : args[arity - 2];
      y = relative ? y + args[arity - 1] : args[arity - 1];
    }

    if (upper === 'M') {
      subpathX = x;
      subpathY = y;
      command = relative ? 'l' : 'L'; // further pairs after a moveto are linetos
    }
    if (start == null) start = { x, y };
    args = [];
  }

  return start == null ? null : { start, end: { x, y } };
}

// Non-negative integer attribute (digits only), or -1.
function parseAnimationIndex(value) {
  return value != null && /^[0-9]{1,6}$/.test(value) ? Number(value) : -1;
}

module.exports = {
  // The slide size converts motion paths, which PowerPoint stores as fractions of the slide,
  // into pixel offsets.
  extractSlideAnimationInfo(slidePart, slideWidthPx = 0, slideHeightPx = 0) {
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

    const shapeBoxes = rawStages.some((stage) => stage.positionAnimationsByTarget.size > 0)
      ? this._collectShapeBoxesPx(slidePart)
      : EmptyShapeBoxes;
    return this._buildSlideAnimationInfo(rawStages, slideWidthPx, slideHeightPx, shapeBoxes);
  },

  _buildSlideAnimationInfo(rawStages, slideWidthPx = 0, slideHeightPx = 0, shapeBoxes = EmptyShapeBoxes) {
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
    // Start offsets of looping position animations, which keep playing in later stages.
    const activePersistentFlyStylesByTarget = new Map();
    // Where each moved target rests after the stages so far (px from its own position).
    const motionOffsetsByTarget = new Map();
    const stages = [];

    for (let index = 0; index < rawStages.length; index++) {
      const rawStage = rawStages[index];
      const rawShapeAnimationCssInfos = this._buildAnimatedShapeCssInfos(rawStage);

      for (const shapeId of rawStage.targetShapeIds) {
        if (rawStage.hideCandidateShapeIds.has(shapeId)) {
          hiddenShapeIds.add(shapeId);
          activePersistentAnimationPartsByShapeId.delete(shapeId);
          activePersistentFlyStylesByTarget.delete(shapeId);
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

      // Motion paths: the target glides from the path's start to its end during this stage
      // and stays at the end point in every later stage (fill="hold").
      const motionStartsByTarget = new Map();
      for (const [target, motionPath] of rawStage.motionPathsByTarget) {
        const start = {
          x: normalizeMotionOffset(motionPath.startX * slideWidthPx),
          y: normalizeMotionOffset(motionPath.startY * slideHeightPx),
        };
        const end = {
          x: normalizeMotionOffset(motionPath.endX * slideWidthPx),
          y: normalizeMotionOffset(motionPath.endY * slideHeightPx),
        };
        motionStartsByTarget.set(target, start);
        if (end.x === 0 && end.y === 0) motionOffsetsByTarget.delete(target);
        else motionOffsetsByTarget.set(target, end);

        const duration = fmtTrim(this.toCssDurationSeconds(motionPath.durationMs, 0.5), 3);
        stageAnimationPartsByShapeId.set(
          target,
          this._mergeAnimationParts(
            [`pptxMotionPath ${duration}s ${motionPath.easing} both`],
            stageAnimationPartsByShapeId.has(target) ? stageAnimationPartsByShapeId.get(target) : null
          )
        );
      }

      // Position animations (Fly In, Float In…): the target glides from its start offset back
      // to its own position during this stage; a looping one keeps going in later stages.
      const flyStylesByTarget = new Map(activePersistentFlyStylesByTarget);
      for (const [target, position] of rawStage.positionAnimationsByTarget) {
        const flyStyle = this._buildFlyInStyle(target, position, shapeBoxes, slideWidthPx, slideHeightPx);
        if (flyStyle == null) continue;

        const duration = fmtTrim(this.toCssDurationSeconds(position.durationMs, 0.5), 3);
        const flyPart = position.isInfiniteLoop
          ? `pptxFlyIn ${duration}s ${position.easing} infinite both`
          : `pptxFlyIn ${duration}s ${position.easing} both`;
        flyStylesByTarget.set(target, flyStyle);
        stageAnimationPartsByShapeId.set(
          target,
          this._mergeAnimationParts(
            [flyPart],
            stageAnimationPartsByShapeId.has(target) ? stageAnimationPartsByShapeId.get(target) : null
          )
        );
        if (position.isInfiniteLoop) {
          activePersistentFlyStylesByTarget.set(target, flyStyle);
          activePersistentAnimationPartsByShapeId.set(
            target,
            this._mergeAnimationParts(
              [flyPart],
              activePersistentAnimationPartsByShapeId.has(target) ? activePersistentAnimationPartsByShapeId.get(target) : null
            )
          );
        }
      }

      stages.push({
        stageNumber: index + 1,
        hiddenShapeIdsAfterStage: this._sortShapeIds(hiddenShapeIds),
        animatedShapeStylesByShapeId: this._buildAnimatedShapeStyles(
          stageAnimationPartsByShapeId,
          motionStartsByTarget,
          motionOffsetsByTarget,
          flyStylesByTarget
        ),
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

    for (const motion of pDescendants(stageElement, 'animMotion')) {
      const ends = parseMotionPathEnds(getAttrValue(motion, 'path'));
      if (ends == null) continue;

      const motionPath = {
        startX: ends.start.x,
        startY: ends.start.y,
        endX: ends.end.x,
        endY: ends.end.y,
        durationMs: this._getDurationMs(pFirstChild(pFirstChild(motion, 'cBhvr'), 'cTn')),
        easing: this._getMotionEasing(motion),
      };
      for (const target of this._getTargetShapeIds(motion)) rawStage.motionPathsByTarget.set(target, motionPath);
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

    for (const animation of pDescendants(stageElement, 'anim')) this._addPositionAnimation(rawStage, animation);

    return rawStage;
  },

  // ParseLinearFormula: { constant, x, y, w, h } or null (see parseLinearFormula above).
  _parseLinearFormula(text) {
    return parseLinearFormula(text);
  },

  // Records a <p:anim> on ppt_x or ppt_y as (one axis of) its targets' position animation.
  _addPositionAnimation(rawStage, animation) {
    const attributeName =
      pDescendants(animation, 'attrName')
        .map((el) => innerText(el))
        .find((name) => name && name.trim() !== '') || '';
    const horizontal = eqIgnoreCase(attributeName, 'ppt_x');
    if (!horizontal && !eqIgnoreCase(attributeName, 'ppt_y')) return;

    const keyframes = this._getPositionKeyframes(animation);
    const durationMs = this._getDurationMs(pFirstChild(pFirstChild(animation, 'cBhvr'), 'cTn'));
    for (const target of this._getTargetShapeIds(animation)) {
      let position = rawStage.positionAnimationsByTarget.get(target);
      if (!position) {
        position = new RawPositionAnimation();
        rawStage.positionAnimationsByTarget.set(target, position);
      }

      if (keyframes == null || (horizontal ? position.fromX : position.fromY) != null) {
        position.isSupported = false;
        continue;
      }

      if (horizontal) {
        position.fromX = keyframes.from;
        position.toX = keyframes.to;
      } else {
        position.fromY = keyframes.from;
        position.toY = keyframes.to;
      }
      position.durationMs = Math.max(position.durationMs, durationMs);
      position.easing = this._getMotionEasing(animation);
      if (this._isIndefiniteRepeat(this._getNearestRepeatCount(animation))) position.isInfiniteLoop = true;
    }
  },

  // Start and end value of a linearly interpolated two-keyframe animation (tm 0 → 100000) whose
  // values are numbers or #ppt_* formulas; null for anything else, such as the time-based fmla
  // keyframes of Bounce.
  _getPositionKeyframes(animation) {
    const calcMode = getAttrValue(animation, 'calcmode') ?? 'lin';
    if (!eqIgnoreCase(calcMode, 'lin')) return null;

    const tavList = pFirstChild(animation, 'tavLst');
    const keyframes = tavList == null ? null : childElements(tavList).filter((el) => isP(el, 'tav'));
    if (
      keyframes == null ||
      keyframes.length !== 2 ||
      (getAttrValue(keyframes[0], 'tm') ?? '0') !== '0' ||
      getAttrValue(keyframes[1], 'tm') !== '100000' ||
      keyframes.some((keyframe) => {
        const formula = getAttrValue(keyframe, 'fmla');
        return formula != null && formula !== '';
      })
    ) {
      return null;
    }

    const from = this._getTimingValueExpression(keyframes[0]);
    const to = this._getTimingValueExpression(keyframes[1]);
    return from.length > 0 && to.length > 0 ? { from, to } : null;
  },

  // The --pptx-fly-from-x/-y start offsets of a position animation, or null when it can't be
  // modelled (unsupported keyframes, or an end point other than the target's own position) or
  // doesn't move. A shape's offsets come from its box on the slide. A paragraph's box is only
  // known once the browser has laid the text out, so its offsets are calc()s over the
  // --pptx-box-* values the box script measures (falling back to the shape's box).
  _buildFlyInStyle(target, position, shapeBoxes, slideWidthPx, slideHeightPx) {
    if (!position.isSupported || slideWidthPx <= 0 || slideHeightPx <= 0) return null;

    const fromX = getStartOffset(position.fromX, position.toX, true, slideWidthPx, slideHeightPx);
    const fromY = getStartOffset(position.fromY, position.toY, false, slideWidthPx, slideHeightPx);
    if (fromX == null || fromY == null || (offsetIsZero(fromX) && offsetIsZero(fromY))) return null;

    const paragraphSeparator = target.indexOf('#');
    const shapeId = paragraphSeparator < 0 ? target : target.substring(0, paragraphSeparator);
    const hasBox = shapeBoxes.has(shapeId);
    const box = hasBox ? shapeBoxes.get(shapeId) : EmptyBox;
    const measured = paragraphSeparator >= 0 || !hasBox;
    const style =
      `--pptx-fly-from-x:${formatFlyOffset(fromX, box, measured)};` +
      `--pptx-fly-from-y:${formatFlyOffset(fromY, box, measured)};`;
    // A paragraph split at line breaks is several elements; the key lets the box script measure
    // them as the one box PowerPoint moves.
    return measured && (offsetDependsOnBox(fromX) || offsetDependsOnBox(fromY))
      ? `--pptx-fly-box:${target};` + style
      : style;
  },

  // Slide-pixel box of every shape on the slide, keyed by shape id, placed the way the shape
  // renderer places it: a placeholder without a size of its own takes its layout's or master's,
  // and group children are mapped through their groups' child spaces.
  _collectShapeBoxesPx(slidePart) {
    const boxes = new Map();
    const shapeTree = shapeTreeOf(slidePart.root);
    if (shapeTree != null) this._collectShapeBoxesIn(shapeTree, slidePart, 0, 0, 0, 0, 1, 1, boxes);
    return boxes;
  },

  // A child box (l, t, w, h) of container lands on the slide at
  // (offsetX + (l - originX) * scaleX, offsetY + (t - originY) * scaleY, w * scaleX, h * scaleY).
  _collectShapeBoxesIn(container, slidePart, offsetX, offsetY, originX, originY, scaleX, scaleY, boxes) {
    const toSlide = ([left, top, width, height]) => ({
      left: offsetX + (left - originX) * scaleX,
      top: offsetY + (top - originY) * scaleY,
      width: width * scaleX,
      height: height * scaleY,
    });

    for (const child of childElements(container)) {
      let shapeId;
      let rect;
      const spPr = shapePropsOf(child);
      if (child.localName === 'sp' && spPr != null) {
        shapeId = this.getPptxShapeId(child);
        rect = this.getTransform(spPr);
        const placeholder = placeholderOf(child);
        if ((rect[2] <= 0 || rect[3] <= 0) && placeholder != null) {
          const resolved = this.resolvePhTransform(placeholder, slidePart);
          if (resolved != null) rect = resolved;
        }
      } else if (child.localName === 'pic' && spPr != null) {
        shapeId = this.getPptxShapeId(child);
        rect = this.getTransform(spPr);
      } else if (child.localName === 'cxnSp' && spPr != null && firstChild(spPr, 'xfrm') != null) {
        shapeId = this.getPptxShapeId(child);
        rect = this.getTransform(spPr);
      } else if (child.localName === 'graphicFrame' && firstChild(child, 'xfrm') != null) {
        const xfrm = firstChild(child, 'xfrm');
        shapeId = this.getPptxShapeId(child);
        rect = [
          this.emuToPx(attrLong(firstChild(xfrm, 'off'), 'x')),
          this.emuToPx(attrLong(firstChild(xfrm, 'off'), 'y')),
          this.emuToPx(attrLong(firstChild(xfrm, 'ext'), 'cx')),
          this.emuToPx(attrLong(firstChild(xfrm, 'ext'), 'cy')),
        ];
      } else if (child.localName === 'grpSp' && firstChild(firstChild(child, 'grpSpPr'), 'xfrm') != null) {
        const xfrm = firstChild(firstChild(child, 'grpSpPr'), 'xfrm');
        const groupBox = toSlide([
          this.emuToPx(attrLong(firstChild(xfrm, 'off'), 'x')),
          this.emuToPx(attrLong(firstChild(xfrm, 'off'), 'y')),
          this.emuToPx(attrLong(firstChild(xfrm, 'ext'), 'cx')),
          this.emuToPx(attrLong(firstChild(xfrm, 'ext'), 'cy')),
        ]);
        const groupId = this.getPptxShapeId(child);
        if (groupId) boxes.set(groupId, groupBox);

        const childSpaceWidth = this.emuToPx(attrLong(firstChild(xfrm, 'chExt'), 'cx'));
        const childSpaceHeight = this.emuToPx(attrLong(firstChild(xfrm, 'chExt'), 'cy'));
        this._collectShapeBoxesIn(
          child,
          slidePart,
          groupBox.left,
          groupBox.top,
          this.emuToPx(attrLong(firstChild(xfrm, 'chOff'), 'x')),
          this.emuToPx(attrLong(firstChild(xfrm, 'chOff'), 'y')),
          groupBox.width > 0 && childSpaceWidth > 0 ? groupBox.width / childSpaceWidth : 1.0,
          groupBox.height > 0 && childSpaceHeight > 0 ? groupBox.height / childSpaceHeight : 1.0,
          boxes
        );
        continue;
      } else {
        continue;
      }

      if (shapeId) boxes.set(shapeId, toSlide(rect));
    }
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

  // flyStylesByTarget: the --pptx-fly-from-* start offsets of the targets a pptxFlyIn plays on.
  _buildAnimatedShapeStyles(
    animationPartsByShapeId,
    motionStartsByTarget = new Map(),
    motionOffsetsByTarget = new Map(),
    flyStylesByTarget = new Map()
  ) {
    if (animationPartsByShapeId.size === 0 && motionOffsetsByTarget.size === 0) return EmptyAnimatedShapeStyles;

    const result = new Map();
    const targets = new Set([...animationPartsByShapeId.keys(), ...motionOffsetsByTarget.keys()]);
    for (const target of targets) {
      const style =
        this._buildMotionStyle(target, motionStartsByTarget, motionOffsetsByTarget) +
        (flyStylesByTarget.has(target) ? flyStylesByTarget.get(target) : '') +
        (animationPartsByShapeId.has(target) ? this._buildShapeAnimationCss(animationPartsByShapeId.get(target)) : '');
      if (style && style.trim() !== '') result.set(target, style);
    }
    return result;
  },

  // The offset a motion path leaves its target at, plus — in the stage that plays the path —
  // the start and end points the pptxMotionPath keyframes glide between.
  _buildMotionStyle(target, motionStartsByTarget, motionOffsetsByTarget) {
    const offset = motionOffsetsByTarget.get(target) || { x: 0, y: 0 };
    let style = '';
    const start = motionStartsByTarget.get(target);
    if (start != null) {
      style +=
        `--pptx-motion-from-x:${f1(start.x)}px;--pptx-motion-from-y:${f1(start.y)}px;` +
        `--pptx-motion-to-x:${f1(offset.x)}px;--pptx-motion-to-y:${f1(offset.y)}px;`;
    }
    if (offset.x !== 0 || offset.y !== 0) style += `transform:translate(${f1(offset.x)}px,${f1(offset.y)}px);`;
    return style;
  },

  // accel/decel sit on the effect's time node (sometimes on the behaviour's own cTn).
  _getMotionEasing(motion) {
    const timeNodes = [pFirstChild(pFirstChild(motion, 'cBhvr'), 'cTn')];
    for (let current = motion.parentNode; current != null && current.nodeType === 1; current = current.parentNode) {
      if (isP(current, 'cTn')) timeNodes.push(current);
    }

    for (const timeNode of timeNodes) {
      const accelerates = parseAnimationIndex(getAttrValue(timeNode, 'accel')) > 0;
      const decelerates = parseAnimationIndex(getAttrValue(timeNode, 'decel')) > 0;
      if (accelerates && decelerates) return 'ease-in-out';
      if (accelerates) return 'ease-in';
      if (decelerates) return 'ease-out';
    }
    return 'linear';
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
  // (<p:set> visibility toggle / <p:animEffect>), a numeric emphasis (<p:anim>) or a motion path
  // (<p:animMotion>). Without one, the emitted sub-slide would be a byte-for-byte duplicate of the
  // main slide (e.g. audio-2-files was producing slideN.html plus an identical slideN_1.html per
  // audio clip).
  _stageHasRenderedEffect(stageElement) {
    return descendants(stageElement).some(
      (element) =>
        isP(element, 'set') || isP(element, 'animEffect') || isP(element, 'anim') || isP(element, 'animMotion')
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

  // Target keys of the effects under element: the shape id, or one "<shape id>#<paragraph>" key
  // per paragraph when the target is a paragraph range (<p:txEl><p:pRg st=".." end=".."/>).
  _getTargetShapeIds(element) {
    const targets = new Set();
    for (const shapeTarget of descendants(element, 'spTgt', PresentationNamespace)) {
      const shapeId = getAttrValue(shapeTarget, 'spid');
      if (!shapeId || shapeId.trim() === '') continue;

      const paragraphRange = pFirstChild(pFirstChild(shapeTarget, 'txEl'), 'pRg');
      const first = parseAnimationIndex(getAttrValue(paragraphRange, 'st'));
      const last = parseAnimationIndex(getAttrValue(paragraphRange, 'end'));
      if (paragraphRange != null && first >= 0 && last >= first && last - first < 1000) {
        for (let paragraph = first; paragraph <= last; paragraph++) targets.add(`${shapeId}#${paragraph}`);
      } else {
        targets.add(shapeId);
      }
    }
    return Array.from(targets);
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
