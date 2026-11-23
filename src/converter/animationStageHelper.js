'use strict';

// Port of PptxToHtml/pptx/AnimationStageHelper.cs

const { childElements, descendants, attr, attrLong, firstChild, innerText, NS } = require('../ooxml/dom');
const { placeholderOf, shapePropsOf, shapeTreeOf } = require('./elements');
const { fmt, fmtTrim, f1 } = require('../util/format');

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
    // CSS timing function of the entrance's accel/decel, and the repeatCount of its effect.
    this.entranceTiming = 'linear';
    this.entranceRepeatCount = '';
    // When the entrance starts, in ms from the start of the click stage.
    this.entranceBeginMs = 0;
    // An exit's clip filter (Wipe, Split, Shape, Wheel, Random Bars… played backwards).
    this.exitFilter = '';
    this.exitDurationMs = 0;
    this.exitBeginMs = 0;
    this.exitTiming = 'linear';
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
    this.trackSegmentsByTarget = new Map();
    this.colorAnimationsByTarget = new Map();
    this.styleSetsByTarget = new Map();
    // Parse-time only: the stage's own time node and the effects met so far, so the segments
    // of one effect share one track effect.
    this.timeNode = null;
    this.trackEffects = new Map();
  }
}

// One behaviour on a target's rotate / scale / opacity / translate / visibility track: an
// animRot or a style.rotation animation, an animScale or a ppt_w / ppt_h animation, an emphasis
// or exit fade, a style.visibility set, one axis of a position animation that the two-keyframe
// Fly In model can't express (Bounce, Fly Out), a curved motion path, a style.opacity set
// (Transparency) or keyframed style.visibility / style.opacity (Blink). beginMs is from the start
// of its effect. Values: degrees (rotate), factors (scale, opacity), 1 / 0 (visible / hidden); a
// formula segment evaluates its keyframes (see evaluateFormulaKeyframes), a path segment its
// points (see evaluateMotionPath).
function trackSegment(fields, timing) {
  return Object.assign(
    {
      property: '',
      kind: '', // by | to | fromTo | fade | formula | set | path | keys
      axis: -1, // 0 = x only, 1 = y only, -1 = both
      fromX: 0,
      fromY: 0,
      toX: 0,
      toY: 0,
      fadeOut: false,
      keyframes: [],
      // keys: values switch at each keyframe (calcmode="discrete") instead of gliding.
      discrete: false,
      // keys: keyframe times (0..1) and values (1 / 0 for visible / hidden, or an opacity).
      keyTimes: [],
      keyValues: [],
      // path: the motion path flattened to a polyline, x0, y0, x1, y1… in fractions of the slide.
      pathPoints: [],
      // Ends on the next click (Transparency "Until Next Click"), so it leaves nothing behind for
      // later stages.
      endsOnNextClick: false,
    },
    fields,
    {
      beginMs: timing.beginMs,
      durationMs: timing.durationMs,
      autoReverse: timing.autoReverse,
      holds: timing.holds,
      accel: timing.accel,
      decel: timing.decel,
      timeFilter: timing.timeFilter,
      effect: timing.effect,
    }
  );
}

// The effect (the time node that carries presetClass) a track segment belongs to: when it starts
// in its click stage, how long one run of it lasts, how often it repeats, and its own
// accel/decel, which remap the time its behaviours see.
const NoTrackEffect = { beginMs: 0, durationMs: 0, repeatCount: '', accel: 0, decel: 0 };

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
    // When the earliest of its animations starts, in ms from the start of the click stage
    // (After Previous / With Previous delays); -1 until one is recorded.
    this.beginMs = -1;
    // False once a ppt_x / ppt_y animation of the target in this stage can't be modelled
    // (time-based fmla, other than two keyframes, a second animation on the same axis, an end
    // point other than the target's own position as in Fly Out…).
    this.isSupported = true;
    // The same animations as translate-track segments, which play them when the Fly In model
    // can't; tracksReadable turns false when one of them can't be read either.
    this.trackSegments = [];
    this.tracksReadable = true;
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

const FORMULA_TOKEN = /#?ppt_[xywh]|(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]+)?|[A-Za-z]+/y;

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
  // Exits write the target's own values without the # (ppt_x): the same thing here.
  switch (token[0].replace(/^#/, '')) {
    case 'ppt_x':
      return linearFormula(0, 1, 0, 0, 0);
    case 'ppt_y':
      return linearFormula(0, 0, 1, 0, 0);
    case 'ppt_w':
      return linearFormula(0, 0, 0, 1, 0);
    case 'ppt_h':
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
// from printing as "-0.0" (which f1 writes, as .NET does).
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

// GetDurationMs: the time node's dur in ms (0 when missing or not a number).
function durationMsOf(timingNode) {
  const n = parseInt(getAttrValue(timingNode, 'dur'), 10);
  return Number.isNaN(n) ? 0 : n;
}

// accel / decel of a time node as a fraction of its duration (100000 = all of it).
function getAccelDecelFraction(timeNode, attributeName) {
  const value = parseAnimationIndex(getAttrValue(timeNode, attributeName));
  return value > 0 ? Math.min(value, 100000) / 100000 : 0;
}

// SMIL accelerate / decelerate, which PowerPoint follows: the speed ramps up linearly over the
// first `accel` of the time, stays constant, and ramps down over the last `decel`. Returns the
// progress (0..1) reached at progress-of-time p.
function applyAccelDecel(p, accel, decel) {
  if ((accel <= 0 && decel <= 0) || accel + decel > 1) return p;

  const rate = 1 / (1 - accel / 2 - decel / 2);
  if (p < accel) return (rate * p * p) / (2 * accel);
  if (p <= 1 - decel) return rate * (p - accel / 2);
  const remaining = 1 - p;
  return 1 - (rate * remaining * remaining) / (2 * decel);
}

const AccelDecelSamples = 10;

// The SMIL accel/decel curve as a CSS linear() easing: exact on the constant-speed part, sampled
// every tenth of each ramp (off by under 0.2 % of the distance).
function buildAccelDecelTimingFunction(accel, decel) {
  if (accel + decel > 1) return 'linear';

  const points = [0];
  for (let step = 1; step <= AccelDecelSamples && accel > 0; step++) points.push((accel * step) / AccelDecelSamples);
  for (let step = 0; step <= AccelDecelSamples && decel > 0; step++) {
    const point = 1 - decel + (decel * step) / AccelDecelSamples;
    if (point > points[points.length - 1] + 1e-9) points.push(point);
  }
  if (points[points.length - 1] < 1 - 1e-9) points.push(1);

  const stops = [];
  for (let index = 0; index < points.length; index++) {
    const point = points[index];
    const value = formatTrackNumber(applyAccelDecel(point, accel, decel), 4);
    stops.push(
      index === 0 || index === points.length - 1 ? value : value + ' ' + formatTrackNumber(point * 100, 2) + '%'
    );
  }
  return 'linear(' + stops.join(', ') + ')';
}

// FormatTrackNumber: up to `digits` decimals, rounded half-to-even like "F" (not fmtTrim, which
// copies the custom "0.###" format and its rounding), and never "-0".
function formatTrackNumber(value, digits) {
  let text = fmt(value, digits);
  if (text.includes('.')) text = text.replace(/0+$/, '').replace(/\.$/, '');
  return text === '-0' ? '0' : text;
}

// ---- Tracks -------------------------------------------------------------------------------
// A track is one animated property of one target in one click stage. Its behaviours are
// evaluated with PowerPoint's timing model (begin delays, accel/decel, autoRev, tmFilter, fill,
// repeatCount) and written out as generated keyframes, sampled at every point where the
// timeline bends. Each property is a separate CSS property (translate, rotate, scale, opacity),
// so tracks combine with each other and with the transform-based keyframes.

const TrackProperties = ['translate', 'rotate', 'scale', 'opacity', 'visibility'];
const TrackSamplesPerSegment = 24;
// How far before / after a jump its extra keyframe goes, as a fraction of the track (0.002 %).
const TrackStepFraction = 2e-5;

const WipeKeyframeRules = [
  ['Down', '@keyframes pptxWipeInDown { from { clip-path:inset(0 0 100% 0); } to { clip-path:inset(0 0 0 0); } }'],
  ['Left', '@keyframes pptxWipeInLeft { from { clip-path:inset(0 0 0 100%); } to { clip-path:inset(0 0 0 0); } }'],
  ['Right', '@keyframes pptxWipeInRight { from { clip-path:inset(0 100% 0 0); } to { clip-path:inset(0 0 0 0); } }'],
  ['Up', '@keyframes pptxWipeInUp { from { clip-path:inset(100% 0 0 0); } to { clip-path:inset(0 0 0 0); } }'],
];

// "wipe(down)" (PowerPoint's Wipe From Top) → "Down": the edge the reveal moves towards.
function getWipeDirection(filter) {
  switch (filter.trim().toLowerCase()) {
    case 'wipe(down)':
      return 'Down';
    case 'wipe(up)':
      return 'Up';
    case 'wipe(left)':
      return 'Left';
    case 'wipe(right)':
      return 'Right';
    default:
      return '';
  }
}

const WipeOutKeyframeRules = [
  ['Down', '@keyframes pptxWipeOutDown { from { clip-path:inset(0 0 0 0); } to { clip-path:inset(0 0 100% 0); } }'],
  ['Left', '@keyframes pptxWipeOutLeft { from { clip-path:inset(0 0 0 0); } to { clip-path:inset(0 0 0 100%); } }'],
  ['Right', '@keyframes pptxWipeOutRight { from { clip-path:inset(0 0 0 0); } to { clip-path:inset(0 100% 0 0); } }'],
  ['Up', '@keyframes pptxWipeOutUp { from { clip-path:inset(0 0 0 0); } to { clip-path:inset(100% 0 0 0); } }'],
];

// ---- Clip filters ------------------------------------------------------------------------------
// Split (barn), Shape (box / circle / diamond / plus), Wheel and Wedge reveal the target through
// a growing clip. Each has a generated keyframes rule: pptxFilterIn<Stem> reveals (entrance),
// pptxFilterOut<Stem> conceals (exit, the same filter played backwards). Geometry, at reveal
// progress q (0 = nothing shown, 1 = all of it), in fractions of the target's box (u right,
// v down, centre .5,.5):
//   barn(outVertical)   |u-.5| <= q/2               barn(inVertical)   |u-.5| >= (1-q)/2
//   barn(outHorizontal) |v-.5| <= q/2               barn(inHorizontal) |v-.5| >= (1-q)/2
//   box(out)  max(|u-.5|,|v-.5|) <= q/2             box(in)  the complement at 1-q
//   circle(out) an ellipse through the corners at q = 1, its radii scaled by q; circle(in) the complement at 1-q
//   diamond(out) |u-.5| + |v-.5| <= q                diamond(in) the complement at 1-q
//   plus(out) a cross of arms q/2 thick reaching q from the centre; plus(in) the complement at 1-q
//   wheel(n) n sectors, each opening clockwise from its spoke (12 o'clock + k*360/n) over q*360/n
//   wedge    one sector opening both ways from 12 o'clock over q*180 each
//   blinds(horizontal) 6 slats across, each opening downwards from its top edge over q of its
//            height; blinds(vertical) the same turned, each opening rightwards
//   checkerboard(across) a grid of 12 x 10 squares, odd rows shifted by one square; each pair of
//            squares opens rightwards over 2q squares, so the board is half open at q = .5;
//            checkerboard(down) the same turned, opening downwards
//   strips(downRight) a staircase of 8 rows sweeping from the top left corner: row j (from that
//            corner) is open over clamp(q*15/8 - j/8, 0, 1) of the width from its side;
//            downLeft / upRight / upLeft start from the other corners
// Wheel and Wedge use true angles (px around the centre), sampled every 15 degrees. The slat and
// square counts are a guess (IE's filters, which PowerPoint's come from, use 10 and 12 x 10):
// measure them against PowerPoint.

const ClipFilters = [
  ['barn(inVertical)', 'BarnInVertical'],
  ['barn(inHorizontal)', 'BarnInHorizontal'],
  ['barn(outVertical)', 'BarnOutVertical'],
  ['barn(outHorizontal)', 'BarnOutHorizontal'],
  ['blinds(horizontal)', 'BlindsHorizontal'],
  ['blinds(vertical)', 'BlindsVertical'],
  ['box(in)', 'BoxIn'],
  ['box(out)', 'BoxOut'],
  ['checkerboard(across)', 'CheckerboardAcross'],
  ['checkerboard(down)', 'CheckerboardDown'],
  ['circle(in)', 'CircleIn'],
  ['circle(out)', 'CircleOut'],
  ['diamond(in)', 'DiamondIn'],
  ['diamond(out)', 'DiamondOut'],
  ['plus(in)', 'PlusIn'],
  ['plus(out)', 'PlusOut'],
  ['strips(downLeft)', 'StripsDownLeft'],
  ['strips(downRight)', 'StripsDownRight'],
  ['strips(upLeft)', 'StripsUpLeft'],
  ['strips(upRight)', 'StripsUpRight'],
  ['wedge', 'Wedge'],
  ['wheel(1)', 'Wheel1'],
  ['wheel(2)', 'Wheel2'],
  ['wheel(3)', 'Wheel3'],
  ['wheel(4)', 'Wheel4'],
  ['wheel(8)', 'Wheel8'],
];

const ClipFilterEllipsePoints = 48;
// Points along each wheel sector's / the wedge's arc (at most 90 degrees apart), on a circle far
// outside any shape, so only the sector's radial edges cut through the target.
const ClipFilterArcPoints = 5;
const ClipFilterArcRadiusPx = 4000;
const ClipFilterAngleStepDegrees = 15;

function getClipFilterStem(filter) {
  const trimmed = filter.trim().toLowerCase();
  for (const [name, stem] of ClipFilters) {
    if (trimmed === name.toLowerCase()) return stem;
  }
  return '';
}

// "randombar(horizontal)" → "horizontal", "randombar(vertical)" → "vertical", else "".
function getRandomBarsOrientation(filter) {
  switch (filter.trim().toLowerCase()) {
    case 'randombar(horizontal)':
      return 'horizontal';
    case 'randombar(vertical)':
      return 'vertical';
    default:
      return '';
  }
}

// True for a filter played as a clip (Wipe, Split, Shape, Wheel, Wedge, Random Bars).
function isClipFilter(filter) {
  return getWipeDirection(filter).length > 0 || getClipFilterStem(filter).length > 0 || getRandomBarsOrientation(filter).length > 0;
}

function formatClipPercent(value) {
  return formatTrackNumber(value, 3) + '%';
}

function formatClipPoint(u, v) {
  return formatClipPercent(u) + ' ' + formatClipPercent(v);
}

// A point at angle degrees (clockwise from 12 o'clock) on the far circle around the centre.
function formatClipArcPoint(degrees) {
  const radians = (degrees * Math.PI) / 180;
  return (
    formatClipCentreOffset(ClipFilterArcRadiusPx * Math.sin(radians)) +
    ' ' +
    formatClipCentreOffset(-ClipFilterArcRadiusPx * Math.cos(radians))
  );
}

function formatClipCentreOffset(offsetPx) {
  const magnitude = formatTrackNumber(Math.abs(offsetPx), 2);
  return magnitude === '0' ? '50%' : `calc(50% ${offsetPx < 0 ? '-' : '+'} ${magnitude}px)`;
}

const ClipBoxOutline = '0% 0%, 100% 0%, 100% 100%, 0% 100%, 0% 0%';

// The clip-path of a clip filter at reveal progress q (see the geometry above).
function getClipFilterShape(stem, q) {
  const inner = 1 - q;
  switch (stem) {
    case 'BarnOutVertical':
      return `inset(0% ${formatClipPercent(50 * inner)} 0% ${formatClipPercent(50 * inner)})`;
    case 'BarnOutHorizontal':
      return `inset(${formatClipPercent(50 * inner)} 0% ${formatClipPercent(50 * inner)} 0%)`;
    case 'BarnInVertical':
      return (
        `polygon(evenodd, ${ClipBoxOutline}, ${formatClipPoint(50 * q, 0)}, ${formatClipPoint(50 * q, 100)}, ` +
        `${formatClipPoint(100 - 50 * q, 100)}, ${formatClipPoint(100 - 50 * q, 0)}, ${formatClipPoint(50 * q, 0)})`
      );
    case 'BarnInHorizontal':
      return (
        `polygon(evenodd, ${ClipBoxOutline}, ${formatClipPoint(0, 50 * q)}, ${formatClipPoint(100, 50 * q)}, ` +
        `${formatClipPoint(100, 100 - 50 * q)}, ${formatClipPoint(0, 100 - 50 * q)}, ${formatClipPoint(0, 50 * q)})`
      );
    case 'BoxOut':
      return `inset(${formatClipPercent(50 * inner)})`;
    case 'BoxIn':
      return (
        `polygon(evenodd, ${ClipBoxOutline}, ${formatClipPoint(50 * q, 50 * q)}, ${formatClipPoint(100 - 50 * q, 50 * q)}, ` +
        `${formatClipPoint(100 - 50 * q, 100 - 50 * q)}, ${formatClipPoint(50 * q, 100 - 50 * q)}, ${formatClipPoint(50 * q, 50 * q)})`
      );
    case 'CircleOut':
      return `ellipse(${formatClipPercent(50 * Math.sqrt(2) * q)} ${formatClipPercent(50 * Math.sqrt(2) * q)} at 50% 50%)`;
    case 'CircleIn': {
      const radius = 50 * Math.sqrt(2) * inner;
      const points = [];
      for (let index = 0; index <= ClipFilterEllipsePoints; index++) {
        const radians = (2 * Math.PI * (index % ClipFilterEllipsePoints)) / ClipFilterEllipsePoints;
        points.push(formatClipPoint(50 + radius * Math.sin(radians), 50 - radius * Math.cos(radians)));
      }
      return `polygon(evenodd, ${ClipBoxOutline}, ${points.join(', ')})`;
    }
    case 'DiamondOut':
      return `polygon(${formatDiamond(100 * q)})`;
    case 'DiamondIn':
      return `polygon(evenodd, ${ClipBoxOutline}, ${formatDiamond(100 * inner)}, ${formatClipPoint(50, 50 - 100 * inner)})`;
    case 'PlusOut':
      return `polygon(${formatPlus(q)})`;
    case 'PlusIn':
      return `polygon(evenodd, ${ClipBoxOutline}, ${formatPlus(inner)}, ${formatClipPoint(50 - 50 * inner, 50 - 100 * inner)})`;
    case 'BlindsHorizontal':
    case 'BlindsVertical':
    case 'CheckerboardAcross':
    case 'CheckerboardDown':
    case 'StripsDownLeft':
    case 'StripsDownRight':
    case 'StripsUpLeft':
    case 'StripsUpRight':
      return formatClipRects(getPatternFilterRects(stem, q));
    case 'Wedge': {
      const points = ['50% 50%'];
      for (let index = 0; index < ClipFilterArcPoints; index++) {
        points.push(formatClipArcPoint(-180 * q + (360 * q * index) / (ClipFilterArcPoints - 1)));
      }
      return `polygon(${points.join(', ')})`;
    }
    default: {
      const spokes = getWheelSpokes(stem);
      const points = [];
      for (let spoke = 0; spoke < spokes; spoke++) {
        points.push('50% 50%');
        const start = (360.0 * spoke) / spokes;
        for (let index = 0; index < ClipFilterArcPoints; index++) {
          points.push(formatClipArcPoint(start + ((360.0 / spokes) * q * index) / (ClipFilterArcPoints - 1)));
        }
      }
      return `polygon(${points.join(', ')})`;
    }
  }
}

const PatternBlindsSlats = 6;
const PatternCheckerColumns = 12;
const PatternCheckerRows = 10;
const PatternStripRows = 8;

// The open rectangles ({ u0, v0, u1, v1 } in percent of the box) of Blinds, Checkerboard or
// Strips at reveal progress q. Their number doesn't depend on q (empty ones stay in), so keyframes
// interpolate point by point; checkerboard squares may reach past the box edge.
function getPatternFilterRects(stem, q) {
  const rects = [];
  const rect = (u0, v0, u1, v1) => rects.push({ u0, v0, u1, v1 });
  if (stem.startsWith('Blinds')) {
    const horizontal = stem === 'BlindsHorizontal';
    for (let slat = 0; slat < PatternBlindsSlats; slat++) {
      const start = (100.0 * slat) / PatternBlindsSlats;
      const end = (100.0 * (slat + q)) / PatternBlindsSlats;
      if (horizontal) rect(0, start, 100, end);
      else rect(start, 0, end, 100);
    }
    return rects;
  }

  if (stem.startsWith('Checkerboard')) {
    const across = stem === 'CheckerboardAcross';
    const along = across ? PatternCheckerColumns : PatternCheckerRows;
    const rows = across ? PatternCheckerRows : PatternCheckerColumns;
    for (let row = 0; row < rows; row++) {
      const rowStart = (100.0 * row) / rows;
      const rowEnd = (100.0 * (row + 1)) / rows;
      const shift = row % 2;
      const pairs = Math.floor((along + shift + 1) / 2);
      for (let pair = 0; pair < pairs; pair++) {
        const start = (100.0 * (2 * pair - shift)) / along;
        const end = (100.0 * (2 * pair - shift + 2 * q)) / along;
        if (across) rect(start, rowStart, end, rowEnd);
        else rect(rowStart, start, rowEnd, end);
      }
    }
    return rects;
  }

  const fromTop = stem.startsWith('StripsDown');
  const fromLeft = stem.endsWith('Right');
  for (let row = 0; row < PatternStripRows; row++) {
    const reach = 100 * Math.min(Math.max((q * (2 * PatternStripRows - 1)) / PatternStripRows - row / PatternStripRows, 0), 1);
    const top = fromTop ? (100.0 * row) / PatternStripRows : (100.0 * (PatternStripRows - row - 1)) / PatternStripRows;
    const bottom = fromTop ? (100.0 * (row + 1)) / PatternStripRows : (100.0 * (PatternStripRows - row)) / PatternStripRows;
    if (fromLeft) rect(0, top, reach, bottom);
    else rect(100 - reach, top, 100, bottom);
  }
  return rects;
}

// Rectangles as one polygon: from a hub at the box's corner out to each rectangle, round it and
// back, so the spokes cancel out and only the rectangles are filled.
function formatClipRects(rects) {
  const points = ['0% 0%'];
  for (const { u0, v0, u1, v1 } of rects) {
    const corner = formatClipPoint(u0, v0);
    points.push(corner, formatClipPoint(u1, v0), formatClipPoint(u1, v1), formatClipPoint(u0, v1), corner, '0% 0%');
  }
  return `polygon(${points.join(', ')})`;
}

// The four corners of a diamond around the centre with half-diagonals of `extent` percent.
function formatDiamond(extent) {
  return (
    `${formatClipPoint(50, 50 - extent)}, ${formatClipPoint(50 + extent, 50)}, ` +
    `${formatClipPoint(50, 50 + extent)}, ${formatClipPoint(50 - extent, 50)}`
  );
}

// The twelve corners of a cross at scale s: arms s/2 thick, reaching s from the centre.
function formatPlus(s) {
  const half = 50 * s;
  const reach = 100 * s;
  return [
    formatClipPoint(50 - half, 50 - reach), formatClipPoint(50 + half, 50 - reach),
    formatClipPoint(50 + half, 50 - half), formatClipPoint(50 + reach, 50 - half),
    formatClipPoint(50 + reach, 50 + half), formatClipPoint(50 + half, 50 + half),
    formatClipPoint(50 + half, 50 + reach), formatClipPoint(50 - half, 50 + reach),
    formatClipPoint(50 - half, 50 + half), formatClipPoint(50 - reach, 50 + half),
    formatClipPoint(50 - reach, 50 - half), formatClipPoint(50 - half, 50 - half),
  ].join(', ');
}

function getWheelSpokes(stem) {
  if (stem.startsWith('Wheel') && /^[0-9]+$/.test(stem.substring(5))) {
    const spokes = parseInt(stem.substring(5), 10);
    if (spokes > 0) return spokes;
  }
  return 1;
}

// Keyframe intervals of a clip filter: the straight-edged ones move linearly (one interval);
// Wheel and Wedge turn, sampled every 15 degrees of each edge's sweep.
function getClipFilterIntervals(stem) {
  // Each strip row starts and stops at a multiple of 1/15 of the time.
  if (stem.startsWith('Strips')) return 2 * PatternStripRows - 1;
  if (stem === 'Wedge') return Math.ceil(180 / ClipFilterAngleStepDegrees);
  if (stem.startsWith('Wheel')) return Math.ceil(360.0 / getWheelSpokes(stem) / ClipFilterAngleStepDegrees);
  return 1;
}

function buildClipFilterRule(name, stem, conceal) {
  const intervals = getClipFilterIntervals(stem);
  const frames = [];
  for (let step = 0; step <= intervals; step++) {
    const time = step / intervals;
    frames.push(`${formatTrackNumber(time * 100, 3)}% { clip-path:${getClipFilterShape(stem, conceal ? 1 - time : time)}; }`);
  }
  return `@keyframes ${name} { ${frames.join(' ')} }`;
}

// Random Bars: N bars across the target, shown one at a time in a fixed shuffled order as
// --pptx-bars runs 0 → 1 (1 → 0 for an exit). The outermost bars reach far beyond the box so
// that nothing (outline, shadow) stays clipped once all bars are shown.
const RandomBarsPropertyRule = "@property --pptx-bars { syntax:'<number>'; inherits:false; initial-value:1; }";
const RandomBarsTargetPx = 4;
const RandomBarsMinCount = 8;
const RandomBarsMaxCount = 64;
const RandomBarsDefaultCount = 32;
const RandomBarsMargin = '1000px';

// The clip-path of a Random Bars target: bars about 4 px thick across its box (32 when the box
// isn't known), each opening once --pptx-bars passes its rank in a shuffled order.
// shuffle(length, seedText) is the converter's buildDeterministicShuffle.
function buildRandomBarsClipStyle(target, orientation, shapeBoxes, shuffle) {
  const horizontal = orientation === 'horizontal';
  const separator = target.indexOf('#');
  const shapeId = separator < 0 ? target : target.substring(0, separator);
  let count = RandomBarsDefaultCount;
  const box = shapeBoxes.get(shapeId);
  if (box != null) {
    const lengthPx = horizontal ? box.height : box.width;
    count = Math.min(Math.max(Math.floor(lengthPx / RandomBarsTargetPx + 0.5), RandomBarsMinCount), RandomBarsMaxCount);
  }

  const order = shuffle(count, `${target}|${orientation}|bars`);
  const ranks = new Array(count).fill(0);
  for (let position = 0; position < count; position++) ranks[order[position]] = position;

  const far = `calc(100% + ${RandomBarsMargin})`;
  const near = '-' + RandomBarsMargin;
  const barSize = formatClipPercent(100.0 / count);
  const points = [];
  for (let bar = 0; bar < count; bar++) {
    const first = bar === 0;
    const last = bar === count - 1;
    const start = first ? near : formatClipPercent((100.0 * bar) / count);
    const span = first || last ? `(${RandomBarsMargin} + ${barSize})` : barSize;
    const shown = `clamp(0, (var(--pptx-bars) - ${formatTrackNumber(ranks[bar] / count, 4)}) * 100000, 1)`;
    const end = `calc(${start} + ${span} * ${shown})`;
    if (horizontal) {
      points.push(`${near} ${start}`, `${far} ${start}`, `${far} ${end}`, `${near} ${end}`);
    } else {
      points.push(`${start} ${near}`, `${start} ${far}`, `${end} ${far}`, `${end} ${near}`);
    }
  }
  return `clip-path:polygon(${points.join(', ')});`;
}

let filterKeyframeRules = null;

// Every keyframes rule a clip filter part may use, by animation name.
function getFilterKeyframeRules() {
  if (filterKeyframeRules != null) return filterKeyframeRules;
  const rules = [];
  for (const [direction, rule] of WipeKeyframeRules) rules.push(['pptxWipeIn' + direction, rule]);
  for (const [direction, rule] of WipeOutKeyframeRules) rules.push(['pptxWipeOut' + direction, rule]);
  for (const [, stem] of ClipFilters) {
    rules.push(['pptxFilterIn' + stem, buildClipFilterRule('pptxFilterIn' + stem, stem, false)]);
    rules.push(['pptxFilterOut' + stem, buildClipFilterRule('pptxFilterOut' + stem, stem, true)]);
  }
  rules.push(['pptxRandomBarsIn', RandomBarsPropertyRule + ' @keyframes pptxRandomBarsIn { from { --pptx-bars:0; } to { --pptx-bars:1; } }']);
  rules.push(['pptxRandomBarsOut', RandomBarsPropertyRule + ' @keyframes pptxRandomBarsOut { from { --pptx-bars:1; } to { --pptx-bars:0; } }']);
  filterKeyframeRules = rules;
  return rules;
}

// ---- Colour changes ----------------------------------------------------------------------------

const ColorChannels = ['fill', 'text', 'stroke'];

// animClr attribute → channel: fillcolor → fill, style.color → text, stroke.color → stroke.
function getColorChannel(attributeName) {
  switch (attributeName.trim().toLowerCase()) {
    case 'fillcolor':
    case 'fill.color':
      return 'fill';
    case 'style.color':
      return 'text';
    case 'stroke.color':
      return 'stroke';
    default:
      return '';
  }
}

function getColorProperty(channel) {
  return channel === 'text' ? 'color' : channel;
}

// The elements a colour channel of a target paints: its filled / outlined geometry, or its text
// (the text box and every paragraph and run, which may set their own colour).
function getColorSelector(target, channel) {
  const scope = `[data-pptx-id="${target}"]`;
  switch (channel) {
    case 'fill':
      return scope + ' :is(rect,path,ellipse,circle,polygon):not([fill="none"])';
    case 'stroke':
      return scope + ' :is(rect,path,ellipse,circle,polygon,polyline,line):not([stroke="none"])';
    default:
      return scope + ' foreignObject :is(div,p,span)';
  }
}

function capitalize(text) {
  return text.charAt(0).toUpperCase() + text.substring(1);
}

// The stage's colour changes as CSS rules: per target and channel, keyframes towards the new
// colour on the elements that channel paints. An animClr with only a "to" starts from each
// element's own colour (or the colour an earlier stage left), so a pulse (autoRev) peaks halfway
// and comes back. A held colour stays on in later stages as a static rule.
// A colour set (Brush Color, dim) switches at its begin and holds the colour for its duration
// (for good when it holds). textStyleParts: each target's text-style animations, which join its
// text colour rule (one element takes one animation list) or get a rule of their own.
function addColorRules(rawStage, colorRestsByTarget, keyframeRulesByName, textStyleParts) {
  const animatedChannels = new Set();
  for (const [target, animations] of rawStage.colorAnimationsByTarget) {
    let rests = colorRestsByTarget.get(target);
    for (const channel of ColorChannels) {
      // The latest-starting animation of a channel plays (two on one channel in one click are
      // rare).
      let animation = null;
      for (const candidate of animations) {
        if (candidate.channel === channel && (animation == null || candidate.beginMs >= animation.beginMs)) animation = candidate;
      }
      if (animation == null) continue;

      animatedChannels.add(target + '|' + channel);
      const property = getColorProperty(channel);
      const from = rests != null && rests.has(channel) ? rests.get(channel) : '';
      const name = 'pptxColor' + capitalize(channel) + sanitizeTrackTarget(target);
      const fromFrame = from.length > 0 ? `0% { ${property}:${from}; } ` : '';
      const frames = animation.isSet
        ? `from { ${property}:${animation.toColor}; } to { ${property}:${animation.toColor}; }`
        : animation.autoReverse
        ? fromFrame + `50% { ${property}:${animation.toColor}; }` + (from.length > 0 ? ` 100% { ${property}:${from}; }` : '')
        : fromFrame + `100% { ${property}:${animation.toColor}; }`;
      const durationMs = animation.isSet
        ? Math.max(animation.durationMs, 1)
        : animation.durationMs * (animation.autoReverse ? 2 : 1);
      let keeps = animation.holds && !animation.autoReverse;
      const fillMode = animation.isSet ? (keeps ? ' forwards' : '') : keeps ? ' both' : ' backwards';
      let part = `${name} ${formatTrackNumber(durationMs / 1000, 3)}s ${animation.timing}${formatAnimationDelay(animation.beginMs)}${fillMode}`;
      if (channel === 'text' && textStyleParts.has(target)) part += ', ' + textStyleParts.get(target).join(', ');
      keyframeRulesByName.set(name, `@keyframes ${name} { ${frames} } ${getColorSelector(target, channel)} { animation:${part}; }`);
      keeps = keeps && !animation.endsOnNextClick;
      if (keeps) {
        if (rests == null) {
          rests = new Map();
          colorRestsByTarget.set(target, rests);
        }
        rests.set(channel, animation.toColor);
      }
    }
  }

  for (const [target, rests] of colorRestsByTarget) {
    for (const [channel, color] of rests) {
      if (animatedChannels.has(target + '|' + channel)) continue;
      const name = 'pptxColorRest' + capitalize(channel) + sanitizeTrackTarget(target);
      keyframeRulesByName.set(name, `${getColorSelector(target, channel)} { ${getColorProperty(channel)}:${color} !important; }`);
    }
  }

  // Text styles of a target whose text colour doesn't change in this stage.
  for (const [target, styleParts] of textStyleParts) {
    if (animatedChannels.has(target + '|text')) continue;
    keyframeRulesByName.set(
      'pptxStyleText' + sanitizeTrackTarget(target),
      `${getColorSelector(target, 'text')} { animation:${styleParts.join(', ')}; }`
    );
  }
}

const TextStyleProperties = ['font-weight', 'font-style', 'text-decoration-line'];

function getTextStyleNameStem(property) {
  switch (property) {
    case 'font-weight':
      return 'FontWeight';
    case 'font-style':
      return 'FontStyle';
    default:
      return 'TextDecoration';
  }
}

// The stage's text-style sets as CSS: per target and property, keyframes that hold the new value
// from the set's begin for its duration (for good when it holds) on the target's text. Returns
// each target's animation parts (addColorRules writes the rule that plays them). Styles held from
// earlier stages stay on as static rules.
function addStyleSetRules(rawStage, styleRestsByTarget, keyframeRulesByName) {
  const partsByTarget = new Map();
  const animatedProperties = new Set();
  for (const [target, sets] of rawStage.styleSetsByTarget) {
    let rests = styleRestsByTarget.get(target);
    for (const property of TextStyleProperties) {
      // The latest-starting set of a property plays.
      let styleSet = null;
      for (const candidate of sets) {
        if (candidate.property === property && (styleSet == null || candidate.beginMs >= styleSet.beginMs)) styleSet = candidate;
      }
      if (styleSet == null) continue;

      animatedProperties.add(target + '|' + property);
      const name = 'pptxStyle' + getTextStyleNameStem(property) + sanitizeTrackTarget(target);
      keyframeRulesByName.set(name, `@keyframes ${name} { from { ${property}:${styleSet.value}; } to { ${property}:${styleSet.value}; } }`);
      if (!partsByTarget.has(target)) partsByTarget.set(target, []);
      const durationMs = Math.max(styleSet.durationMs, 1);
      partsByTarget
        .get(target)
        .push(`${name} ${formatTrackNumber(durationMs / 1000, 3)}s linear${formatAnimationDelay(styleSet.beginMs)}${styleSet.holds ? ' forwards' : ''}`);
      if (styleSet.holds && !styleSet.endsOnNextClick) {
        if (rests == null) {
          rests = new Map();
          styleRestsByTarget.set(target, rests);
        }
        rests.set(property, styleSet.value);
      }
    }
  }

  for (const [target, rests] of styleRestsByTarget) {
    for (const [property, value] of rests) {
      if (animatedProperties.has(target + '|' + property)) continue;
      const name = 'pptxStyleRest' + getTextStyleNameStem(property) + sanitizeTrackTarget(target);
      keyframeRulesByName.set(name, `${getColorSelector(target, 'text')} { ${property}:${value} !important; }`);
    }
  }
  return partsByTarget;
}

// ---- Exits and delays --------------------------------------------------------------------------

// " 0.5s": a CSS animation-delay for an effect starting beginMs into its click stage, or "".
function formatAnimationDelay(beginMs) {
  return beginMs > 0 ? ' ' + formatTrackNumber(beginMs / 1000, 3) + 's' : '';
}

// True when a target's exit in this stage has something to play before the target is gone (a
// clip, a fade, a move, a resize…, or a visibility switch later than the stage's start), so it
// stays rendered on the stage's page. Disappear, hidden at once, has nothing.
function isAnimatedExit(rawStage, target) {
  const shapeAnimation = rawStage.shapeAnimationsByShapeId.get(target);
  if (shapeAnimation != null && shapeAnimation.exitFilter.length > 0) return true;
  if (rawStage.positionAnimationsByTarget.has(target) || rawStage.motionPathsByTarget.has(target)) return true;
  const segments = rawStage.trackSegmentsByTarget.get(target);
  return segments != null && segments.some((segment) => segment.property !== 'visibility' || segment.effect.beginMs + segment.beginMs > 0);
}

// ---- Style sets, keyframed styles, motion paths ----------------------------------------------

// The CSS property and value of a text-style set (style.fontWeight bold → font-weight:bold,
// style.fontStyle italic, style.textDecorationUnderline true → underline), or ['', ''].
function getTextStyleSet(attributeName, value) {
  const text = value.trim().toLowerCase();
  switch (attributeName.trim().toLowerCase()) {
    case 'style.fontweight':
      return text === 'bold' || text === 'normal' ? ['font-weight', text] : ['', ''];
    case 'style.fontstyle':
      return text === 'italic' || text === 'normal' || text === 'oblique' ? ['font-style', text] : ['', ''];
    case 'style.textdecorationunderline':
      if (text === 'true' || text === '1') return ['text-decoration-line', 'underline'];
      return text === 'false' || text === '0' ? ['text-decoration-line', 'none'] : ['', ''];
    default:
      return ['', ''];
  }
}

// The value of a <p:to> (or a <p:tav>'s <p:val>): its fltVal / strVal / intVal / boolVal.
function getAnimationValueText(value) {
  for (const name of ['fltVal', 'strVal', 'intVal', 'boolVal']) {
    const text = getAttrValue(pFirstChild(value, name), 'val');
    if (text != null) return text;
  }
  return '';
}

// Whether a style set keeps its value once its own duration is over: fill="hold" or "freeze", or
// no end at all (dur="indefinite": Until End of Slide). A timed set with no fill is removed when
// it ends (SMIL's fill="auto"), as a timed Transparency is.
function styleSetHolds(behaviour) {
  const timeNode = pFirstChild(pFirstChild(behaviour, 'cBhvr'), 'cTn');
  const fill = getAttrValue(timeNode, 'fill');
  if (eqIgnoreCase(fill, 'hold') || eqIgnoreCase(fill, 'freeze')) return true;
  if (eqIgnoreCase(fill, 'remove')) return false;
  return durationMsOf(timeNode) <= 0;
}

// True when the behaviour's time node, or one above it inside its click stage, ends on the next
// click (<p:endCondLst><p:cond evt="onNext"/>): Transparency… Until Next Click.
function endsOnNextClick(rawStage, behaviour) {
  const timeNodes = [pFirstChild(pFirstChild(behaviour, 'cBhvr'), 'cTn')];
  for (
    let current = behaviour.parentNode;
    current != null && current.nodeType === 1 && current !== rawStage.timeNode;
    current = current.parentNode
  ) {
    if (isP(current, 'cTn')) timeNodes.push(current);
  }

  for (const timeNode of timeNodes) {
    const conditions = pFirstChild(timeNode, 'endCondLst');
    if (
      conditions != null &&
      childElements(conditions).some((condition) => isP(condition, 'cond') && eqIgnoreCase(getAttrValue(condition, 'evt'), 'onNext'))
    ) {
      return true;
    }
  }
  return false;
}

// A keys segment's value at progress: the keyframe in force (discrete) or the value interpolated
// between the two keyframes around it.
function evaluateStyleKeyframes(segment, progress) {
  const times = segment.keyTimes;
  const values = segment.keyValues;
  let index = 0;
  while (index + 1 < times.length && progress >= times[index + 1]) index++;
  if (segment.discrete || index + 1 >= times.length || progress <= times[index]) return values[index];

  const span = times[index + 1] - times[index];
  return span > 0 ? values[index] + ((values[index + 1] - values[index]) * (progress - times[index])) / span : values[index + 1];
}

const MotionPathCurveSteps = 16;

// A motion path that isn't a single straight line as a polyline (x0, y0, x1, y1… in fractions of
// the slide, from the target's own position), each cubic curve cut into 16 straight pieces and a
// Z closing back to its subpath's start; empty for one straight line or a point (pptxMotionPath
// plays those).
function flattenMotionPath(pathText) {
  if (pathText == null || pathText.trim() === '') return [];

  const points = [];
  let x = 0;
  let y = 0;
  let subpathX = 0;
  let subpathY = 0;
  let command = ' ';
  let pieces = 0;
  let curved = false;
  let args = [];
  for (const token of pathText.match(MOTION_PATH_TOKEN) || []) {
    if (/^[A-Za-z]$/.test(token)) {
      command = token;
      args = [];
      if (command === 'Z' || command === 'z') {
        x = subpathX;
        y = subpathY;
        points.push(x, y);
        pieces++;
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
    const originX = relative ? x : 0;
    const originY = relative ? y : 0;
    if (upper === 'C') {
      const x1 = originX + args[0];
      const y1 = originY + args[1];
      const x2 = originX + args[2];
      const y2 = originY + args[3];
      const x3 = originX + args[4];
      const y3 = originY + args[5];
      for (let step = 1; step <= MotionPathCurveSteps; step++) {
        const t = step / MotionPathCurveSteps;
        const u = 1 - t;
        const a = u * u * u;
        const b = 3 * u * u * t;
        const c = 3 * u * t * t;
        const d = t * t * t;
        points.push(a * x + b * x1 + c * x2 + d * x3, a * y + b * y1 + c * y2 + d * y3);
      }
      x = x3;
      y = y3;
      curved = true;
      pieces++;
    } else {
      if (upper === 'H') x = originX + args[0];
      else if (upper === 'V') y = originY + args[0];
      else {
        x = originX + args[0];
        y = originY + args[1];
      }

      if (upper === 'M') {
        subpathX = x;
        subpathY = y;
        command = relative ? 'l' : 'L'; // further pairs after a moveto are linetos
      } else {
        pieces++;
      }
      points.push(x, y);
    }
    args = [];
  }

  return curved || pieces > 1 ? points : [];
}

// Distance (px) along a flattened path to each of its points.
function getMotionPathLengths(points, slideWidthPx, slideHeightPx) {
  const count = Math.floor(points.length / 2);
  const lengths = new Array(count).fill(0);
  let total = 0;
  for (let index = 1; index < count; index++) {
    const dx = (points[2 * index] - points[2 * index - 2]) * slideWidthPx;
    const dy = (points[2 * index + 1] - points[2 * index - 1]) * slideHeightPx;
    total += Math.sqrt(dx * dx + dy * dy);
    lengths[index] = total;
  }
  return lengths;
}

// Where a path segment puts its target at progress (px from its own position): PowerPoint moves
// along the path at a steady speed, so progress is a share of the path's length.
function evaluateMotionPath(points, progress, context) {
  const count = Math.floor(points.length / 2);
  if (count === 0) return { x: 0, y: 0 };

  const lengths = getMotionPathLengths(points, context.slideWidthPx, context.slideHeightPx);
  const target = lengths[count - 1] * Math.min(Math.max(progress, 0), 1);
  let index = 1;
  while (index < count - 1 && lengths[index] < target) index++;
  if (count === 1) return { x: points[0] * context.slideWidthPx, y: points[1] * context.slideHeightPx };

  const span = lengths[index] - lengths[index - 1];
  const fraction = span > 0 ? (target - lengths[index - 1]) / span : 1;
  const x = points[2 * index - 2] + (points[2 * index] - points[2 * index - 2]) * fraction;
  const y = points[2 * index - 1] + (points[2 * index + 1] - points[2 * index - 1]) * fraction;
  return { x: x * context.slideWidthPx, y: y * context.slideHeightPx };
}

// The time fraction at which applyAccelDecel reaches progress y (its inverse).
function invertAccelDecel(y, accel, decel) {
  if ((accel <= 0 && decel <= 0) || accel + decel > 1) return y;

  const rate = 1 / (1 - accel / 2 - decel / 2);
  if (y < (rate * accel) / 2) return Math.sqrt((2 * accel * y) / rate);
  if (y <= 1 - (rate * decel) / 2) return y / rate + accel / 2;
  return 1 - Math.sqrt(Math.max(0, (2 * decel * (1 - y)) / rate));
}

// The times (ms from the track's start) at which a path segment reaches each corner of its
// polyline, through its own and its effect's accel/decel, so the keyframes trace the path.
function addMotionPathSampleTimes(times, segment, effectStart, context) {
  const lengths = getMotionPathLengths(segment.pathPoints, context.slideWidthPx, context.slideHeightPx);
  const total = lengths.length > 0 ? lengths[lengths.length - 1] : 0;
  if (!(total > 0) || segment.durationMs <= 0) return;

  const effect = segment.effect;
  const passes = segment.autoReverse ? 2 : 1;
  for (const length of lengths) {
    const local = invertAccelDecel(length / total, segment.accel, segment.decel);
    for (let pass = 0; pass < passes; pass++) {
      let effectMs = segment.beginMs + segment.durationMs * (pass === 0 ? local : 2 - local);
      if (isRemappedEffect(effect)) {
        const fraction = effectMs / effect.durationMs;
        if (fraction > 1) continue;
        effectMs = effect.durationMs * invertAccelDecel(fraction, effect.accel, effect.decel);
      }
      times.push(effectStart + effectMs);
    }
  }
}

// True when a position formula puts the target back at its own position along the axis (#ppt_x,
// or ppt_x as exits write it).
function endsAtOwnPosition(to, horizontal) {
  const value = parseLinearFormula(to);
  return (
    value != null &&
    Math.abs(value.constant) < 1e-12 &&
    value.w === 0 &&
    value.h === 0 &&
    (horizontal ? value.x === 1 && value.y === 0 : value.x === 0 && value.y === 1)
  );
}

// True when the nearest effect node above element is an emphasis (presetClass="emph").
function isEmphasisEffect(element) {
  for (let current = element.parentNode; current != null && current.nodeType === 1; current = current.parentNode) {
    if (!isP(current, 'cTn')) continue;
    const presetClass = getAttrValue(current, 'presetClass');
    if (presetClass != null) return eqIgnoreCase(presetClass, 'emph');
  }
  return false;
}

const SIGNED_INTEGER = /^[-+]?[0-9]{1,10}$/;
const TRACK_NUMBER = /^[-+]?(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]+)?$/;

// An animRot angle (60000ths of a degree) in degrees, or null.
function parseAnimationAngle(value) {
  return value != null && SIGNED_INTEGER.test(value) ? Number(value) / 60000 : null;
}

// An animScale <p:by>/<p:from>/<p:to> (x, y in 1000ths of a percent) as factors, or null.
function parseScaleVector(element) {
  const x = getAttrValue(element, 'x');
  const y = getAttrValue(element, 'y');
  if (x == null || y == null || !SIGNED_INTEGER.test(x) || !SIGNED_INTEGER.test(y)) return null;
  return { x: Number(x) / 100000, y: Number(y) / 100000 };
}

function tryParseTrackNumber(text) {
  const trimmed = text.trim();
  return TRACK_NUMBER.test(trimmed) ? Number(trimmed) : null;
}

function isTrueAttribute(value) {
  return value === '1' || eqIgnoreCase(value, 'true');
}

// The numeric start delay of a time node (its first <p:cond delay="…">), or 0.
function getStartDelayMs(timeNode) {
  const conditions = pFirstChild(timeNode, 'stCondLst');
  if (conditions == null) return 0;
  for (const condition of childElements(conditions).filter((el) => isP(el, 'cond'))) {
    const delay = parseAnimationIndex(getAttrValue(condition, 'delay'));
    if (delay >= 0) return delay;
  }
  return 0;
}

// One run of an effect: its own dur, else where its last behaviour ends.
function getEffectDurationMs(effectNode) {
  const explicitMs = durationMsOf(effectNode);
  if (explicitMs > 0) return explicitMs;

  let endMs = 0;
  for (const timeNode of pDescendants(effectNode, 'cTn').filter((el) => isP(el.parentNode, 'cBhvr'))) {
    let beginMs = 0;
    for (let current = timeNode; current != null && current.nodeType === 1 && current !== effectNode; current = current.parentNode) {
      if (isP(current, 'cTn')) beginMs += getStartDelayMs(current);
    }
    const passes = isTrueAttribute(getAttrValue(timeNode, 'autoRev')) ? 2 : 1;
    endMs = Math.max(endMs, beginMs + durationMsOf(timeNode) * passes);
  }
  return endMs;
}

// tmFilter="0,0; .2,.5; 1,0": (time, progress) pairs that remap a behaviour's progress.
function parseTimeFilter(text) {
  if (text == null || text.trim() === '') return [];

  const values = [];
  for (const pair of text.split(';')) {
    if (pair.trim().length === 0) continue;
    const parts = pair.split(',');
    const time = parts.length === 2 ? tryParseTrackNumber(parts[0]) : null;
    const value = parts.length === 2 ? tryParseTrackNumber(parts[1]) : null;
    if (time == null || value == null) return [];
    values.push(time, value);
  }
  return values.length >= 4 ? values : [];
}

function applyTimeFilter(filter, progress) {
  if (filter.length < 4) return progress;
  if (progress <= filter[0]) return filter[1];

  for (let index = 2; index + 1 < filter.length; index += 2) {
    if (progress > filter[index]) continue;
    const span = filter[index] - filter[index - 2];
    return span > 0
      ? filter[index - 1] + ((filter[index + 1] - filter[index - 1]) * (progress - filter[index - 2])) / span
      : filter[index + 1];
  }
  return filter[filter.length - 1];
}

// k for a size formula k*#ppt_w (k*#ppt_h when !width); null for anything else.
function tryGetSizeFactor(formula, width) {
  const factor = width ? formula.w : formula.h;
  return formula.x === 0 && formula.y === 0 && (width ? formula.h : formula.w) === 0 && Math.abs(formula.constant) < 1e-12
    ? factor
    : null;
}

function getTimingValueText(timingValue) {
  const valueElement = pFirstChild(timingValue, 'val');
  for (const name of ['fltVal', 'strVal', 'intVal']) {
    const value = getAttrValue(pFirstChild(valueElement, name), 'val');
    if (value != null) return value;
  }
  return '';
}

function getTrackFormulaContext(target, shapeBoxes, slideWidthPx, slideHeightPx) {
  const separator = target.indexOf('#');
  const shapeId = separator < 0 ? target : target.substring(0, separator);
  if (slideWidthPx <= 0 || slideHeightPx <= 0 || !shapeBoxes.has(shapeId)) {
    return { x: NaN, y: NaN, w: NaN, h: NaN, slideWidthPx, slideHeightPx };
  }
  const box = shapeBoxes.get(shapeId);
  return {
    x: (box.left + box.width / 2) / slideWidthPx,
    y: (box.top + box.height / 2) / slideHeightPx,
    w: box.width / slideWidthPx,
    h: box.height / slideHeightPx,
    slideWidthPx,
    slideHeightPx,
  };
}

function getTrackIdentity(property) {
  return property === 'scale' || property === 'opacity' || property === 'visibility' ? { x: 1, y: 1 } : { x: 0, y: 0 };
}

function isTrackIdentity(property, value) {
  return formatTrackDeclaration(property, value) === formatTrackDeclaration(property, getTrackIdentity(property));
}

function formatTrackDeclaration(property, value) {
  switch (property) {
    case 'rotate':
      return `rotate:${formatTrackNumber(value.x, 3)}deg;`;
    case 'scale':
      return `scale:${formatTrackNumber(value.x, 4)} ${formatTrackNumber(value.y, 4)};`;
    case 'opacity':
      return `opacity:${formatTrackNumber(value.x, 4)};`;
    case 'visibility':
      return value.x >= 0.5 ? 'visibility:visible;' : 'visibility:hidden;';
    default:
      return `translate:${formatFlyPx(value.x)} ${formatFlyPx(value.y)};`;
  }
}

// The values a target holds after earlier stages' tracks, as static declarations (with the
// transform box when no animation of the stage sets it).
function buildTrackRestStyle(rests, includeTransformBox) {
  let declarations = '';
  for (const property of TrackProperties) {
    if (rests.has(property)) declarations += formatTrackDeclaration(property, rests.get(property));
  }
  if (declarations.length === 0) return '';
  return (includeTransformBox ? 'transform-box:fill-box;transform-origin:center center;' : '') + declarations;
}

function isRepeatingEffect(effect) {
  const count = parseAnimationIndex(effect.repeatCount);
  return eqIgnoreCase(effect.repeatCount, 'indefinite') || (count > 0 && count !== 1000);
}

function getTrackIterations(effect) {
  return eqIgnoreCase(effect.repeatCount, 'indefinite')
    ? 'infinite'
    : formatTrackNumber(parseAnimationIndex(effect.repeatCount) / 1000, 3);
}

function isRemappedEffect(effect) {
  return effect.durationMs > 0 && (effect.accel > 0 || effect.decel > 0);
}

function getSegmentEndMs(segment) {
  const spanMs = segment.durationMs * (segment.autoReverse ? 2 : 1);
  return (
    segment.effect.beginMs +
    Math.max(segment.beginMs + spanMs, isRemappedEffect(segment.effect) ? segment.effect.durationMs : 0)
  );
}

function isNonLinearSegment(segment) {
  return (
    segment.accel > 0 ||
    segment.decel > 0 ||
    (segment.kind === 'formula' &&
      (segment.keyframes.length > 2 || segment.keyframes.some((keyframe) => keyframe.formula.length > 0)))
  );
}

function sanitizeTrackTarget(target) {
  return target.replace(/[^0-9A-Za-z-]/g, '_');
}

// The generated keyframes of one property of one target (segments sorted by start), or null
// when it can't be played: effects that repeat mixed with others, nothing to time, a formula
// that can't be evaluated. A repeating effect gives a looping animation of one run.
function buildTrackAnimation(target, property, segments, baseValue, context) {
  const effects = [...new Set(segments.map((segment) => segment.effect))];
  const loops = effects.some(isRepeatingEffect);
  if (loops && effects.length > 1) return null;

  const startMs = loops ? effects[0].beginMs : 0;
  let lengthMs = loops ? effects[0].durationMs : Math.max(...segments.map(getSegmentEndMs));
  // Sets with no duration (Transparency until the next click or the end of the slide) still need
  // a timeline to hold their value: a 1 ms one.
  if (!loops && lengthMs <= 0 && segments.every((segment) => segment.kind === 'set')) lengthMs = 1;
  if (!(lengthMs > 0)) return null;

  const name = 'pptxTrack' + property[0].toUpperCase() + property.substring(1) + sanitizeTrackTarget(target);
  const times = getTrackSampleTimes(segments, startMs, lengthMs, context);
  const { starts, removals } = getTrackStepTimes(segments, startMs);
  const stepMs = lengthMs * TrackStepFraction;
  const frames = [];
  let value = baseValue;
  let lastTime = 0;
  let lastValue = baseValue;
  for (let index = 0; index < times.length; index++) {
    const time = times[index];
    value = evaluateTrack(segments, startMs + time, baseValue, context);
    if (!isFiniteTrackValue(value)) return null;

    // SMIL holds a value up to the instant a later behaviour starts (or a fill="remove" one
    // ends). Where the value jumps there, a keyframe just before (after) the jump keeps CSS from
    // interpolating across it; it is left out where CSS gets it right anyway.
    if (index > 0 && time - 2 * stepMs >= times[index - 1] && isTrackStepTime(starts, time)) {
      const before = evaluateTrack(segments, startMs + time - stepMs, baseValue, context);
      const interpolated = lerpTrackValue(lastValue, value, (time - stepMs - lastTime) / (time - lastTime));
      if (isFiniteTrackValue(before) && !sameTrackDeclaration(property, before, interpolated)) {
        frames.push(formatTrackFrame(time - stepMs, lengthMs, formatTrackDeclaration(property, before)));
      }
    }

    // A behaviour removed (fill="remove") right at the end of the track: the value it leaves
    // behind holds from the last keyframe on, so it jumps there.
    if (!loops && index === times.length - 1 && index > 0 && time - 2 * stepMs >= times[index - 1] && isTrackStepTime(removals, time)) {
      const after = evaluateTrack(segments, Infinity, baseValue, context);
      if (isFiniteTrackValue(after) && !sameTrackDeclaration(property, after, value)) {
        frames.push(formatTrackFrame(time - stepMs, lengthMs, formatTrackDeclaration(property, value)));
        value = after;
      }
    }

    frames.push(formatTrackFrame(time, lengthMs, formatTrackDeclaration(property, value)));
    lastTime = time;
    lastValue = value;

    if (index + 1 < times.length && time + 2 * stepMs <= times[index + 1] && isTrackStepTime(removals, time)) {
      const nextTime = times[index + 1];
      const next = evaluateTrack(segments, startMs + nextTime, baseValue, context);
      const after = evaluateTrack(segments, startMs + time + stepMs, baseValue, context);
      const interpolated = lerpTrackValue(value, next, stepMs / (nextTime - time));
      if (isFiniteTrackValue(after) && !sameTrackDeclaration(property, after, interpolated)) {
        frames.push(formatTrackFrame(time + stepMs, lengthMs, formatTrackDeclaration(property, after)));
        lastTime = time + stepMs;
        lastValue = after;
      }
    }
  }

  // A visibility track that never changes (an entrance shown at once) has nothing to play.
  if (property === 'visibility' && new Set(frames.map((frame) => frame.substring(frame.indexOf('{')))).size <= 1) return null;

  const duration = formatTrackNumber(lengthMs / 1000, 3);
  const part = loops
    ? `${name} ${duration}s linear${startMs > 0 ? ' ' + formatTrackNumber(startMs / 1000, 3) + 's' : ''} ${getTrackIterations(effects[0])} both`
    : `${name} ${duration}s linear both`;
  return {
    name,
    keyframesRule: `@keyframes ${name} { ${frames.join(' ')} }`,
    animationPart: part,
    loops,
    endValue: value,
  };
}

function isFiniteTrackValue(value) {
  return Number.isFinite(value.x) && Number.isFinite(value.y);
}

function lerpTrackValue(from, to, fraction) {
  return { x: from.x + (to.x - from.x) * fraction, y: from.y + (to.y - from.y) * fraction };
}

function sameTrackDeclaration(property, a, b) {
  return formatTrackDeclaration(property, a) === formatTrackDeclaration(property, b);
}

function formatTrackFrame(time, lengthMs, declaration) {
  return formatTrackNumber((time / lengthMs) * 100, 3) + '% { ' + declaration + ' }';
}

function isTrackStepTime(stepTimes, time) {
  return stepTimes.some((stepTime) => Math.abs(stepTime - time) < 1e-6);
}

// Where a behaviour can make its track jump (ms from startMs): where it starts, since it
// replaces what earlier behaviours left, and, with fill="remove", where it ends. Effects with
// accel/decel are skipped (their time is remapped and densely sampled anyway).
function getTrackStepTimes(segments, startMs) {
  const starts = [];
  const removals = [];
  for (const segment of segments) {
    if (isRemappedEffect(segment.effect)) continue;
    const begin = segment.effect.beginMs - startMs + segment.beginMs;
    starts.push(begin);
    if (!segment.holds) removals.push(begin + segment.durationMs * (segment.autoReverse ? 2 : 1));
    // A discrete keyframe (Blink) switches the value at its own time too.
    if (segment.kind === 'keys' && segment.discrete) addKeyframeTimes(starts, segment, begin);
  }
  return { starts, removals };
}

// The times of a keys segment's keyframes in every pass (autoRev plays them backwards).
function addKeyframeTimes(times, segment, begin) {
  const passes = segment.autoReverse ? 2 : 1;
  for (const keyTime of segment.keyTimes) {
    for (let pass = 0; pass < passes; pass++) times.push(begin + segment.durationMs * (pass === 0 ? keyTime : 2 - keyTime));
  }
}

// Times (ms from startMs, within 0..lengthMs) where the track's timeline bends: every
// behaviour's start and end, its tmFilter points, and evenly spaced samples through curved
// stretches (accel/decel, formulas).
function getTrackSampleTimes(segments, startMs, lengthMs, context) {
  const times = [0, lengthMs];
  for (const segment of segments) {
    const effect = segment.effect;
    const effectStart = effect.beginMs - startMs;
    // A path's corners and a keyframed style's keyframes are where its timeline bends.
    if (segment.kind === 'path') addMotionPathSampleTimes(times, segment, effectStart, context);
    else if (segment.kind === 'keys' && !isRemappedEffect(effect)) addKeyframeTimes(times, segment, effectStart + segment.beginMs);
    if (isRemappedEffect(effect)) {
      for (let step = 0; step <= TrackSamplesPerSegment; step++) {
        times.push(effectStart + (effect.durationMs * step) / TrackSamplesPerSegment);
      }
      continue;
    }

    const begin = effectStart + segment.beginMs;
    const passes = segment.autoReverse ? 2 : 1;
    const curved = isNonLinearSegment(segment);
    times.push(begin);
    for (let pass = 0; pass < passes; pass++) {
      const passStart = begin + segment.durationMs * pass;
      times.push(passStart + segment.durationMs);
      if (curved) {
        for (let step = 1; step < TrackSamplesPerSegment; step++) {
          times.push(passStart + (segment.durationMs * step) / TrackSamplesPerSegment);
        }
      }
      for (let index = 0; index + 1 < segment.timeFilter.length; index += 2) {
        const point = segment.timeFilter[index];
        times.push(passStart + segment.durationMs * (pass === 0 ? point : 1 - point));
      }
    }
  }

  times.sort((a, b) => a - b);
  const result = [];
  for (const time of times) {
    if (time < 0 || time > lengthMs || (result.length > 0 && time - result[result.length - 1] < 1e-6)) continue;
    result.push(time);
  }
  return result;
}

// The track's value at stageMs: the segments apply in start order on top of baseValue (what
// earlier stages left), each from its begin and, with fill="hold", after its end.
function evaluateTrack(segments, stageMs, baseValue, context) {
  let value = baseValue;
  for (const segment of segments) {
    const effect = segment.effect;
    let effectMs = stageMs - effect.beginMs;
    if (effectMs < 0) continue;
    if (isRemappedEffect(effect)) {
      effectMs = effect.durationMs * applyAccelDecel(Math.min(effectMs / effect.durationMs, 1), effect.accel, effect.decel);
    }

    const localMs = effectMs - segment.beginMs;
    if (localMs < 0) continue;
    const spanMs = segment.durationMs * (segment.autoReverse ? 2 : 1);
    if (localMs > spanMs && !segment.holds) continue;

    value = applyTrackSegment(segment, value, getSegmentProgress(segment, Math.min(localMs, spanMs)), context);
  }
  return value;
}

function getSegmentProgress(segment, localMs) {
  let progress = 1;
  if (segment.durationMs > 0) {
    progress = localMs / segment.durationMs;
    if (progress > 1) progress = segment.autoReverse ? Math.max(0, 2 - progress) : 1;
  }
  return applyTimeFilter(segment.timeFilter, applyAccelDecel(progress, segment.accel, segment.decel));
}

function applyTrackSegment(segment, value, progress, context) {
  const applyX = segment.axis !== 1;
  const applyY = segment.axis !== 0;
  switch (segment.kind) {
    case 'by':
      if (segment.property === 'rotate') return { x: value.x + segment.toX * progress, y: value.y };
      return {
        x: applyX ? value.x * (1 + (segment.toX - 1) * progress) : value.x,
        y: applyY ? value.y * (1 + (segment.toY - 1) * progress) : value.y,
      };
    case 'to':
      return {
        x: applyX ? value.x + (segment.toX - value.x) * progress : value.x,
        y: applyY ? value.y + (segment.toY - value.y) * progress : value.y,
      };
    case 'fromTo':
      return {
        x: applyX ? segment.fromX + (segment.toX - segment.fromX) * progress : value.x,
        y: applyY ? segment.fromY + (segment.toY - segment.fromY) * progress : value.y,
      };
    case 'fade':
      return { x: value.x * (segment.fadeOut ? 1 - progress : progress), y: value.y };
    case 'set':
      return { x: segment.toX, y: segment.toY };
    case 'formula': {
      const offset = evaluateFormulaKeyframes(segment, progress, context);
      return segment.axis === 0 ? { x: offset, y: value.y } : { x: value.x, y: offset };
    }
    case 'path':
      return evaluateMotionPath(segment.pathPoints, progress, context);
    case 'keys': {
      const keyed = evaluateStyleKeyframes(segment, progress);
      return { x: keyed, y: keyed };
    }
    default:
      return value;
  }
}

// A position animation's offset (px from the target's own position) at progress: the value
// interpolated between the two keyframes around it ($), passed through the earlier keyframe's
// fmla when it has one ("#ppt_y-sin(pi*$)/3").
function evaluateFormulaKeyframes(segment, progress, context) {
  const keyframes = segment.keyframes;
  let index = 0;
  while (index < keyframes.length - 2 && progress > keyframes[index + 1].time) index++;

  const start = keyframes[index];
  const end = keyframes.length > 1 ? keyframes[index + 1] : start;
  const startValue = evaluateFormula(start.value, context, NaN);
  const endValue = evaluateFormula(end.value, context, NaN);
  const span = end.time - start.time;
  const local = span > 0 ? Math.min(Math.max((progress - start.time) / span, 0), 1) : 1;
  const dollar = startValue + (endValue - startValue) * local;
  const value = start.formula.length > 0 ? evaluateFormula(start.formula, context, dollar) : dollar;
  return segment.axis === 0
    ? (value - context.x) * context.slideWidthPx
    : (value - context.y) * context.slideHeightPx;
}

const TRACK_FORMULA_TOKEN = /#?ppt_[xywh]|(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][-+]?[0-9]+)?|[A-Za-z]+|\$/y;

// Evaluates a timing formula ("#ppt_y-sin(pi*$)/3"): numbers, #ppt_x/#ppt_y/#ppt_w/#ppt_h, $
// (dollar), pi, e, + - * / ^, parentheses and PowerPoint's functions. NaN when it can't.
function evaluateFormula(text, context, dollar) {
  const state = { text, position: 0, context, dollar };
  const value = evaluateFormulaSum(state);
  skipFormulaSpaces(state);
  return state.position === text.length ? value : NaN;
}

function evaluateFormulaSum(state) {
  let left = evaluateFormulaProduct(state);
  while (!Number.isNaN(left)) {
    skipFormulaSpaces(state);
    const c = state.text[state.position];
    if (state.position >= state.text.length || (c !== '+' && c !== '-')) break;
    state.position++;
    const right = evaluateFormulaProduct(state);
    left = c === '-' ? left - right : left + right;
  }
  return left;
}

function evaluateFormulaProduct(state) {
  let left = evaluateFormulaPower(state);
  while (!Number.isNaN(left)) {
    skipFormulaSpaces(state);
    const c = state.text[state.position];
    if (state.position >= state.text.length || (c !== '*' && c !== '/')) break;
    state.position++;
    const right = evaluateFormulaPower(state);
    left = c === '/' ? left / right : left * right;
  }
  return left;
}

function evaluateFormulaPower(state) {
  const left = evaluateFormulaUnary(state);
  skipFormulaSpaces(state);
  if (Number.isNaN(left) || state.position >= state.text.length || state.text[state.position] !== '^') return left;
  state.position++;
  return Math.pow(left, evaluateFormulaPower(state));
}

function evaluateFormulaUnary(state) {
  skipFormulaSpaces(state);
  const text = state.text;
  if (state.position < text.length && (text[state.position] === '-' || text[state.position] === '+')) {
    const negate = text[state.position++] === '-';
    const operand = evaluateFormulaUnary(state);
    return negate ? -operand : operand;
  }

  if (state.position < text.length && text[state.position] === '(') {
    state.position++;
    const inner = evaluateFormulaSum(state);
    skipFormulaSpaces(state);
    if (state.position >= text.length || text[state.position] !== ')') return NaN;
    state.position++;
    return inner;
  }

  TRACK_FORMULA_TOKEN.lastIndex = state.position;
  const token = TRACK_FORMULA_TOKEN.exec(text);
  if (token == null) return NaN;
  state.position += token[0].length;
  const name = token[0].toLowerCase().replace(/^#/, '');
  switch (name) {
    case 'ppt_x':
      return state.context.x;
    case 'ppt_y':
      return state.context.y;
    case 'ppt_w':
      return state.context.w;
    case 'ppt_h':
      return state.context.h;
    case '$':
      return state.dollar;
    case 'pi':
      return Math.PI;
    case 'e':
      return Math.E;
  }

  if ((name[0] >= '0' && name[0] <= '9') || name[0] === '.') return Number(name);

  skipFormulaSpaces(state);
  if (state.position >= text.length || text[state.position] !== '(') return NaN;
  state.position++;
  const args = [evaluateFormulaSum(state)];
  skipFormulaSpaces(state);
  while (state.position < text.length && text[state.position] === ',') {
    state.position++;
    args.push(evaluateFormulaSum(state));
    skipFormulaSpaces(state);
  }
  if (state.position >= text.length || text[state.position] !== ')') return NaN;
  state.position++;

  const x = args[0];
  if (args.length === 2) {
    if (name === 'max') return Math.max(x, args[1]);
    if (name === 'min') return Math.min(x, args[1]);
    return NaN;
  }
  if (args.length !== 1) return NaN;
  switch (name) {
    case 'abs':
      return Math.abs(x);
    case 'acos':
      return Math.acos(x);
    case 'asin':
      return Math.asin(x);
    case 'atan':
      return Math.atan(x);
    case 'ceil':
      return Math.ceil(x);
    case 'cos':
      return Math.cos(x);
    case 'cosh':
      return Math.cosh(x);
    case 'deg':
      return (x * 180) / Math.PI;
    case 'exp':
      return Math.exp(x);
    case 'floor':
      return Math.floor(x);
    case 'ln':
      return Math.log(x);
    case 'rad':
      return (x * Math.PI) / 180;
    case 'sin':
      return Math.sin(x);
    case 'sinh':
      return Math.sinh(x);
    case 'sqrt':
      return Math.sqrt(x);
    case 'tan':
      return Math.tan(x);
    case 'tanh':
      return Math.tanh(x);
    default:
      return NaN;
  }
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

    // Shape boxes place position animations and size Random Bars' bars.
    const shapeBoxes = rawStages.some(
      (stage) =>
        stage.positionAnimationsByTarget.size > 0 ||
        [...stage.shapeAnimationsByShapeId.values()].some(
          (info) => getRandomBarsOrientation(info.entranceFilter).length > 0 || getRandomBarsOrientation(info.exitFilter).length > 0
        )
    )
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
    // Track values a target holds after the stages so far (fill="hold"), by property, and the
    // keyframes of looping tracks, which keep playing in later stages.
    const trackRestsByTarget = new Map();
    const persistentKeyframeRulesByTarget = new Map();
    // Colours a target keeps after the stages so far (Fill Color, Font Color… with fill="hold"),
    // by channel (fill / text / stroke).
    const colorRestsByTarget = new Map();
    // Text styles a target keeps after the stages so far (Bold Reveal, Underline…), by CSS property.
    const styleRestsByTarget = new Map();
    const stages = [];

    for (let index = 0; index < rawStages.length; index++) {
      const rawStage = rawStages[index];
      const rawShapeAnimationCssInfos = this._buildAnimatedShapeCssInfos(rawStage, shapeBoxes);
      const hiddenBeforeStage = new Set(hiddenShapeIds);
      // Targets whose exit animates in this stage: they stay rendered on this stage's page while
      // the exit plays, and are hidden from the next stage on.
      const animatedExitTargets = [];

      for (const shapeId of rawStage.targetShapeIds) {
        if (rawStage.hideCandidateShapeIds.has(shapeId)) {
          hiddenShapeIds.add(shapeId);
          if (isAnimatedExit(rawStage, shapeId)) {
            animatedExitTargets.push(shapeId);
            continue;
          }

          activePersistentAnimationPartsByShapeId.delete(shapeId);
          activePersistentFlyStylesByTarget.delete(shapeId);
          trackRestsByTarget.delete(shapeId);
          persistentKeyframeRulesByTarget.delete(shapeId);
          colorRestsByTarget.delete(shapeId);
          styleRestsByTarget.delete(shapeId);
          continue;
        }
        if (rawStage.showCandidateShapeIds.has(shapeId)) hiddenShapeIds.delete(shapeId);
      }

      const stageAnimationPartsByShapeId = new Map();
      for (const [k, v] of activePersistentAnimationPartsByShapeId) {
        stageAnimationPartsByShapeId.set(k, v.slice());
      }

      // Static declarations the stage's effects need (Random Bars' clip-path).
      const effectStylesByTarget = new Map();
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
        if (cssInfo.style.length > 0) effectStylesByTarget.set(shapeId, cssInfo.style);
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
        // Path points are offsets from the target's own position, so where an earlier curved path
        // left it no longer counts.
        if (trackRestsByTarget.has(target)) trackRestsByTarget.get(target).delete('translate');

        const duration = fmtTrim(this.toCssDurationSeconds(motionPath.durationMs, 0.5), 3);
        stageAnimationPartsByShapeId.set(
          target,
          this._mergeAnimationParts(
            [`pptxMotionPath ${duration}s ${motionPath.easing}${formatAnimationDelay(motionPath.beginMs)} both`],
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
        // A loop carries on in later stages without its delay (it is already running).
        const flyPart = position.isInfiniteLoop
          ? `pptxFlyIn ${duration}s ${position.easing} infinite both`
          : `pptxFlyIn ${duration}s ${position.easing} both`;
        const stageFlyPart = position.isInfiniteLoop
          ? `pptxFlyIn ${duration}s ${position.easing}${formatAnimationDelay(position.beginMs)} infinite both`
          : `pptxFlyIn ${duration}s ${position.easing}${formatAnimationDelay(position.beginMs)} both`;
        flyStylesByTarget.set(target, flyStyle);
        stageAnimationPartsByShapeId.set(
          target,
          this._mergeAnimationParts(
            [stageFlyPart],
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

      for (const [target, effectStyle] of effectStylesByTarget) {
        flyStylesByTarget.set(target, (flyStylesByTarget.has(target) ? flyStylesByTarget.get(target) : '') + effectStyle);
      }

      // Tracks (rotation, scale, emphasis and exit fades, formula positions, visibility): each
      // property of a target plays its own generated keyframes, starting from what earlier
      // stages left.
      const keyframeRulesByName = new Map();
      for (const rules of persistentKeyframeRulesByTarget.values()) {
        for (const [name, rule] of rules) keyframeRulesByName.set(name, rule);
      }

      for (const [target, segments] of this._collectTrackSegments(rawStage)) {
        // A curved path's points are offsets from the target's own position too.
        if (segments.some((segment) => segment.kind === 'path')) motionOffsetsByTarget.delete(target);
        let rests = trackRestsByTarget.has(target) ? trackRestsByTarget.get(target) : null;
        const context = getTrackFormulaContext(target, shapeBoxes, slideWidthPx, slideHeightPx);
        const trackParts = [];
        const loopingTrackParts = [];
        for (const property of TrackProperties) {
          const propertySegments = segments
            .filter((segment) => segment.property === property)
            .sort((a, b) => a.effect.beginMs + a.beginMs - (b.effect.beginMs + b.beginMs));
          if (propertySegments.length === 0) continue;

          // Visibility starts from whether the target was shown before this stage; the hidden sets
          // carry it into later stages, so it never rests.
          const visibility = property === 'visibility';
          const baseValue = visibility
            ? hiddenBeforeStage.has(target)
              ? { x: 0.0, y: 0.0 }
              : { x: 1.0, y: 1.0 }
            : rests != null && rests.has(property)
              ? rests.get(property)
              : getTrackIdentity(property);
          const track = buildTrackAnimation(target, property, propertySegments, baseValue, context);
          if (track == null) continue;

          keyframeRulesByName.set(track.name, track.keyframesRule);
          trackParts.push(track.animationPart);
          if (visibility) continue;
          if (track.loops) {
            if (!persistentKeyframeRulesByTarget.has(target)) persistentKeyframeRulesByTarget.set(target, new Map());
            persistentKeyframeRulesByTarget.get(target).set(track.name, track.keyframesRule);
            loopingTrackParts.push(track.animationPart);
            continue;
          }

          // What lasts past the next click: without the behaviours that end on it.
          const endValue = propertySegments.some((segment) => segment.endsOnNextClick)
            ? evaluateTrack(
                propertySegments.filter((segment) => !segment.endsOnNextClick),
                Infinity,
                baseValue,
                context
              )
            : track.endValue;
          if (isTrackIdentity(property, endValue)) {
            if (rests != null) rests.delete(property);
            continue;
          }

          if (rests == null) {
            rests = new Map();
            trackRestsByTarget.set(target, rests);
          }
          rests.set(property, endValue);
        }

        if (loopingTrackParts.length > 0) {
          activePersistentAnimationPartsByShapeId.set(
            target,
            this._mergeAnimationParts(
              loopingTrackParts,
              activePersistentAnimationPartsByShapeId.has(target) ? activePersistentAnimationPartsByShapeId.get(target) : null
            )
          );
        }

        if (trackParts.length > 0) {
          stageAnimationPartsByShapeId.set(
            target,
            this._mergeAnimationParts(
              trackParts,
              stageAnimationPartsByShapeId.has(target) ? stageAnimationPartsByShapeId.get(target) : null
            )
          );
        }
      }

      // Text styles and colour changes: rules on the target's text / fill / outline elements; what
      // earlier stages left as static rules.
      const textStyleParts = addStyleSetRules(rawStage, styleRestsByTarget, keyframeRulesByName);
      addColorRules(rawStage, colorRestsByTarget, keyframeRulesByName, textStyleParts);

      for (const [name, rule] of getFilterKeyframeRules()) {
        const used = [...stageAnimationPartsByShapeId.values()].some((parts) =>
          parts.some((part) => part.startsWith(name + ' '))
        );
        if (used) keyframeRulesByName.set(name, rule);
      }

      const restStylesByTarget = new Map();
      for (const [target, rests] of trackRestsByTarget) {
        const restStyle = buildTrackRestStyle(rests, !stageAnimationPartsByShapeId.has(target));
        if (restStyle.length > 0) restStylesByTarget.set(target, restStyle);
      }

      // An exit's target leaves nothing to carry into later stages, where it is hidden.
      for (const target of animatedExitTargets) {
        activePersistentAnimationPartsByShapeId.delete(target);
        activePersistentFlyStylesByTarget.delete(target);
        trackRestsByTarget.delete(target);
        persistentKeyframeRulesByTarget.delete(target);
        colorRestsByTarget.delete(target);
        styleRestsByTarget.delete(target);
      }

      stages.push({
        stageNumber: index + 1,
        hiddenShapeIdsAfterStage: this._sortShapeIds([...hiddenShapeIds].filter((shapeId) => !animatedExitTargets.includes(shapeId))),
        animatedShapeStylesByShapeId: this._buildAnimatedShapeStyles(
          stageAnimationPartsByShapeId,
          motionStartsByTarget,
          motionOffsetsByTarget,
          flyStylesByTarget,
          restStylesByTarget
        ),
        keyframeRules: [...keyframeRulesByName.entries()]
          .sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0))
          .map((entry) => entry[1]),
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

    rawStage.timeNode = commonTimeNode;
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
      if (!targetsVisibility) {
        this._addStyleSet(rawStage, setElement);
        continue;
      }

      const strVal = pDescendants(setElement, 'strVal')[0];
      const value = getAttrValue(strVal, 'val');
      const visible = eqIgnoreCase(value, 'visible');
      if (!visible && !eqIgnoreCase(value, 'hidden')) continue;

      // An emphasis (Blink) switches visibility for a while; it doesn't show or hide the target
      // for good, so only its visibility track sees it.
      if (!isEmphasisEffect(setElement)) {
        for (const id of this._getTargetShapeIds(setElement)) {
          if (visible) rawStage.showCandidateShapeIds.add(id);
          else rawStage.hideCandidateShapeIds.add(id);
        }
      }

      // When it happens: an After Previous entrance stays hidden until its turn, an exit stays
      // visible until its effect ends.
      this._addTrackSegments(rawStage, setElement, (timing) =>
        trackSegment(
          { property: 'visibility', kind: 'set', toX: visible ? 1 : 0, toY: visible ? 1 : 0 },
          { ...timing, autoReverse: false, accel: 0, decel: 0, timeFilter: [] }
        )
      );
    }

    for (const animationEffect of pDescendants(stageElement, 'animEffect')) {
      const transition = getAttrValue(animationEffect, 'transition');
      const filter = getAttrValue(animationEffect, 'filter') || '';
      const durationMs = this._getDurationMs(pFirstChild(pFirstChild(animationEffect, 'cBhvr'), 'cTn'));
      const transitionOut = eqIgnoreCase(transition, 'out');

      // An emphasis (Pulse, Flash Bulb…) fades the shape out part way and back; it never shows
      // or hides it, so it plays on the opacity track instead.
      if (isEmphasisEffect(animationEffect)) {
        if (filter.toLowerCase().includes('fade')) this._addFadeTrackSegments(rawStage, animationEffect, transitionOut);
        continue;
      }

      const timing = this._getTrackTiming(rawStage, animationEffect);
      if (eqIgnoreCase(transition, 'in')) {
        for (const id of this._getTargetShapeIds(animationEffect)) rawStage.showCandidateShapeIds.add(id);
      } else if (transitionOut) {
        // An exit plays its filter backwards: a fade (and any filter without a clip model) on the
        // opacity track, Wipe / Split / Shape / Wheel / Random Bars as a clip.
        for (const id of this._getTargetShapeIds(animationEffect)) rawStage.hideCandidateShapeIds.add(id);
        if (!isClipFilter(filter)) {
          this._addFadeTrackSegments(rawStage, animationEffect, true);
          continue;
        }

        for (const shapeId of this._getTargetShapeIds(animationEffect)) {
          const shapeAnimation = this._getOrCreateShapeAnimationInfo(rawStage, shapeId);
          shapeAnimation.exitFilter = filter.trim();
          shapeAnimation.exitDurationMs = durationMs;
          shapeAnimation.exitBeginMs = timing.effect.beginMs + timing.beginMs;
          shapeAnimation.exitTiming = this._getMotionEasing(animationEffect);
        }
        continue;
      }

      for (const shapeId of this._getTargetShapeIds(animationEffect)) {
        if (!eqIgnoreCase(transition, 'in')) continue;
        const shapeAnimation = this._getOrCreateShapeAnimationInfo(rawStage, shapeId);
        shapeAnimation.entranceTransition = transition || '';
        if (filter && filter.trim() !== '') shapeAnimation.entranceFilter = filter;
        if (durationMs > 0) shapeAnimation.entranceDurationMs = durationMs;
        shapeAnimation.entranceTiming = this._getMotionEasing(animationEffect);
        shapeAnimation.entranceRepeatCount = this._getNearestRepeatCount(animationEffect);
        shapeAnimation.entranceBeginMs = timing.effect.beginMs + timing.beginMs;
      }
    }

    // Colour changes (Color Pulse, Fill Color, Font Color, Line Color): <p:animClr> to a colour,
    // on a shape's fillcolor, style.color (its text) or stroke.color.
    for (const colorAnimation of pDescendants(stageElement, 'animClr')) {
      const channel = getColorChannel(
        pDescendants(colorAnimation, 'attrName')
          .map((el) => innerText(el))
          .find((name) => name && name.trim() !== '') || ''
      );
      const toColor = pFirstChild(colorAnimation, 'to');
      if (
        channel.length === 0 ||
        toColor == null ||
        (firstChild(toColor, 'srgbClr', NS.A) == null && firstChild(toColor, 'schemeClr', NS.A) == null)
      ) {
        continue;
      }

      const timing = this._getTrackTiming(rawStage, colorAnimation);
      const animation = {
        channel,
        toColor: '#' + this.resolveColorElement(toColor),
        beginMs: timing.effect.beginMs + timing.beginMs,
        durationMs: timing.durationMs,
        autoReverse: timing.autoReverse,
        holds: timing.holds,
        timing: this._getMotionEasing(colorAnimation),
      };
      for (const target of this._getTargetShapeIds(colorAnimation)) {
        // Paragraph targets (Font Color by paragraph) aren't modelled.
        if (target.includes('#')) continue;
        if (!rawStage.colorAnimationsByTarget.has(target)) rawStage.colorAnimationsByTarget.set(target, []);
        rawStage.colorAnimationsByTarget.get(target).push(animation);
      }
    }

    // Rotations (Spin, Teeter): animRot by / from / to, in 60000ths of a degree.
    for (const rotation of pDescendants(stageElement, 'animRot')) {
      const by = parseAnimationAngle(getAttrValue(rotation, 'by'));
      const from = parseAnimationAngle(getAttrValue(rotation, 'from'));
      const to = parseAnimationAngle(getAttrValue(rotation, 'to'));
      if (by == null && to == null) continue;

      this._addTrackSegments(rawStage, rotation, (timing) =>
        trackSegment(
          {
            property: 'rotate',
            kind: by != null ? 'by' : from != null ? 'fromTo' : 'to',
            axis: 0,
            fromX: from ?? 0,
            toX: by ?? to ?? 0,
          },
          timing
        )
      );
    }

    // Scales (Pulse, Grow/Shrink, Bounce's squash): animScale <p:by>/<p:from>/<p:to> with x/y in
    // 1000ths of a percent.
    for (const scale of pDescendants(stageElement, 'animScale')) {
      const by = parseScaleVector(pFirstChild(scale, 'by'));
      const from = parseScaleVector(pFirstChild(scale, 'from'));
      const to = parseScaleVector(pFirstChild(scale, 'to'));
      if (by == null && to == null) continue;

      const end = by ?? to;
      this._addTrackSegments(rawStage, scale, (timing) =>
        trackSegment(
          {
            property: 'scale',
            kind: by != null ? 'by' : from != null ? 'fromTo' : 'to',
            fromX: from != null ? from.x : 1,
            fromY: from != null ? from.y : 1,
            toX: end.x,
            toY: end.y,
          },
          timing
        )
      );
    }

    for (const motion of pDescendants(stageElement, 'animMotion')) {
      const pathText = getAttrValue(motion, 'path');
      const ends = parseMotionPathEnds(pathText);
      if (ends == null) continue;

      // A path that isn't one straight line (arcs, turns, loops, closed shapes) plays on the
      // translate track, point by point; a straight one glides with pptxMotionPath.
      const pathPoints = flattenMotionPath(pathText);
      if (pathPoints.length > 0) {
        this._addTrackSegments(rawStage, motion, (timing) => trackSegment({ property: 'translate', kind: 'path', pathPoints }, timing));
        continue;
      }

      const motionTiming = this._getTrackTiming(rawStage, motion);
      const motionPath = {
        startX: ends.start.x,
        startY: ends.start.y,
        endX: ends.end.x,
        endY: ends.end.y,
        durationMs: this._getDurationMs(pFirstChild(pFirstChild(motion, 'cBhvr'), 'cTn')),
        easing: this._getMotionEasing(motion),
        beginMs: motionTiming.effect.beginMs + motionTiming.beginMs,
      };
      for (const target of this._getTargetShapeIds(motion)) rawStage.motionPathsByTarget.set(target, motionPath);
    }

    for (const animation of pDescendants(stageElement, 'anim')) {
      const attributeName =
        pDescendants(animation, 'attrName')
          .map((el) => innerText(el))
          .find((name) => name && name.trim() !== '') || '';
      if (attributeName.trim() === '') continue;

      // Keyframed visibility or opacity (Blink): the visibility / opacity track.
      if (this._addStyleKeyframeTrack(rawStage, animation, attributeName)) continue;

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
      const isTrack = this._addSizeOrRotationTrack(rawStage, animation, attributeName);

      for (const shapeId of this._getTargetShapeIds(animation)) {
        const shapeAnimation = this._getOrCreateShapeAnimationInfo(rawStage, shapeId);
        shapeAnimation.numericAnimationsByAttributeName.set(attributeName.toLowerCase(), {
          durationMs,
          repeatCount,
          expressions,
          changesValue,
          hasSineExpression,
          isTrack,
        });
      }
    }

    for (const animation of pDescendants(stageElement, 'anim')) this._addPositionAnimation(rawStage, animation);

    return rawStage;
  },

  // A <p:set> other than on visibility: style.opacity (Transparency) on the opacity track, a
  // colour (Brush Color, After Animation dim) as a colour rule, or a text style (Bold Flash, Bold
  // Reveal, Underline) as a rule on the target's text. Paragraph targets aren't modelled.
  _addStyleSet(rawStage, setElement) {
    const attributeName =
      pDescendants(setElement, 'attrName')
        .map((el) => innerText(el))
        .find((name) => name && name.trim() !== '') || '';
    const to = pFirstChild(setElement, 'to');
    if (to == null) return;

    const holds = styleSetHolds(setElement);
    const endsOnNext = endsOnNextClick(rawStage, setElement);
    if (eqIgnoreCase(attributeName.trim(), 'style.opacity')) {
      let opacity = tryParseTrackNumber(getAnimationValueText(to));
      if (opacity == null) return;
      opacity = Math.min(Math.max(opacity, 0), 1);
      this._addTrackSegments(rawStage, setElement, (timing) =>
        trackSegment(
          { property: 'opacity', kind: 'set', toX: opacity, toY: opacity, endsOnNextClick: endsOnNext },
          { ...timing, holds, autoReverse: false, accel: 0, decel: 0, timeFilter: [] }
        )
      );
      return;
    }

    const timing = this._getTrackTiming(rawStage, setElement);
    const beginMs = timing.effect.beginMs + timing.beginMs;
    const channel = getColorChannel(attributeName);
    if (channel.length > 0) {
      const colorValue = pFirstChild(to, 'clrVal');
      if (colorValue == null || (firstChild(colorValue, 'srgbClr', NS.A) == null && firstChild(colorValue, 'schemeClr', NS.A) == null)) return;

      const animation = {
        channel,
        toColor: '#' + this.resolveColorElement(colorValue),
        beginMs,
        durationMs: timing.durationMs,
        autoReverse: false,
        holds,
        timing: 'linear',
        isSet: true,
        endsOnNextClick: endsOnNext,
      };
      for (const target of this._getTargetShapeIds(setElement)) {
        if (target.includes('#')) continue;
        if (!rawStage.colorAnimationsByTarget.has(target)) rawStage.colorAnimationsByTarget.set(target, []);
        rawStage.colorAnimationsByTarget.get(target).push(animation);
      }
      return;
    }

    const [property, value] = getTextStyleSet(attributeName, getAnimationValueText(to));
    if (property.length === 0) return;

    const styleSet = { property, value, beginMs, durationMs: timing.durationMs, holds, endsOnNextClick: endsOnNext };
    for (const target of this._getTargetShapeIds(setElement)) {
      if (target.includes('#')) continue;
      if (!rawStage.styleSetsByTarget.has(target)) rawStage.styleSetsByTarget.set(target, []);
      rawStage.styleSetsByTarget.get(target).push(styleSet);
    }
  },

  // A <p:anim> with keyframes on style.visibility (Blink: hidden for the first half, then visible)
  // or style.opacity: a visibility / opacity track whose values switch (discrete, always for
  // visibility) or glide between the keyframes. True when it was taken.
  _addStyleKeyframeTrack(rawStage, animation, attributeName) {
    const visibility = eqIgnoreCase(attributeName.trim(), 'style.visibility');
    if (!visibility && !eqIgnoreCase(attributeName.trim(), 'style.opacity')) return false;

    const tavList = pFirstChild(animation, 'tavLst');
    const keyframes = tavList == null ? null : childElements(tavList).filter((el) => isP(el, 'tav'));
    if (keyframes == null || keyframes.length === 0) return false;

    const times = [];
    const values = [];
    for (let index = 0; index < keyframes.length; index++) {
      const time = parseAnimationIndex(getAttrValue(keyframes[index], 'tm'));
      times.push(time >= 0 ? time / 100000 : keyframes.length > 1 ? index / (keyframes.length - 1) : 0);
      const text = getTimingValueText(keyframes[index]).trim();
      if (visibility) {
        if (eqIgnoreCase(text, 'visible')) values.push(1);
        else if (eqIgnoreCase(text, 'hidden')) values.push(0);
        else return false;
      } else {
        const opacity = tryParseTrackNumber(text);
        if (opacity == null) return false;
        values.push(Math.min(Math.max(opacity, 0), 1));
      }
    }

    const discrete = visibility || eqIgnoreCase(getAttrValue(animation, 'calcmode'), 'discrete');
    this._addTrackSegments(rawStage, animation, (timing) =>
      trackSegment(
        { property: visibility ? 'visibility' : 'opacity', kind: 'keys', discrete, keyTimes: times, keyValues: values },
        timing
      )
    );
    return true;
  },

  // FormatTrackNumber (see formatTrackNumber above).
  _formatTrackNumber(value, digits) {
    return formatTrackNumber(value, digits);
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

    let keyframes = this._getPositionKeyframes(animation);
    // The Fly In model glides back to the target's own position; anything that ends elsewhere
    // (Fly Out and the other exits) plays on the translate track instead.
    if (keyframes != null && !endsAtOwnPosition(keyframes.to, horizontal)) keyframes = null;
    const durationMs = this._getDurationMs(pFirstChild(pFirstChild(animation, 'cBhvr'), 'cTn'));
    const formulaKeyframes = this._readFormulaKeyframes(animation);
    const timing = this._getTrackTiming(rawStage, animation);
    const beginMs = timing.effect.beginMs + timing.beginMs;
    for (const target of this._getTargetShapeIds(animation)) {
      let position = rawStage.positionAnimationsByTarget.get(target);
      if (!position) {
        position = new RawPositionAnimation();
        rawStage.positionAnimationsByTarget.set(target, position);
      }
      if (position.beginMs < 0 || beginMs < position.beginMs) position.beginMs = beginMs;

      if (formulaKeyframes == null) {
        position.tracksReadable = false;
      } else {
        position.trackSegments.push(
          trackSegment(
            { property: 'translate', kind: 'formula', axis: horizontal ? 0 : 1, keyframes: formulaKeyframes },
            timing
          )
        );
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

  _buildAnimatedShapeCssInfos(rawStage, shapeBoxes = EmptyShapeBoxes) {
    const animationCssInfosByShapeId = new Map();
    for (const [shapeId, shapeAnimationInfo] of rawStage.shapeAnimationsByShapeId) {
      const animationCssInfo = this._buildShapeAnimationCssInfo(shapeAnimationInfo, shapeId, shapeBoxes);
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
    flyStylesByTarget = new Map(),
    restStylesByTarget = new Map()
  ) {
    if (animationPartsByShapeId.size === 0 && motionOffsetsByTarget.size === 0 && restStylesByTarget.size === 0) {
      return EmptyAnimatedShapeStyles;
    }

    const result = new Map();
    const targets = new Set([
      ...animationPartsByShapeId.keys(),
      ...motionOffsetsByTarget.keys(),
      ...restStylesByTarget.keys(),
    ]);
    for (const target of targets) {
      const style =
        this._buildMotionStyle(target, motionStartsByTarget, motionOffsetsByTarget) +
        (flyStylesByTarget.has(target) ? flyStylesByTarget.get(target) : '') +
        (restStylesByTarget.has(target) ? restStylesByTarget.get(target) : '') +
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
      const accel = getAccelDecelFraction(timeNode, 'accel');
      const decel = getAccelDecelFraction(timeNode, 'decel');
      if (accel > 0 || decel > 0) return buildAccelDecelTimingFunction(accel, decel);
    }
    return 'linear';
  },

  // The timing of one behaviour: its begin and duration (ms, from its effect's start), autoRev,
  // fill, accel/decel, tmFilter, and its effect.
  _getTrackTiming(rawStage, behaviour) {
    const timeNode = pFirstChild(pFirstChild(behaviour, 'cBhvr'), 'cTn');
    const ancestors = [];
    for (
      let current = behaviour.parentNode;
      current != null && current.nodeType === 1 && current !== rawStage.timeNode;
      current = current.parentNode
    ) {
      if (isP(current, 'cTn')) ancestors.push(current);
    }

    let effectIndex = ancestors.findIndex((node) => getAttrValue(node, 'presetClass') != null);
    if (effectIndex < 0) effectIndex = ancestors.length > 0 ? 0 : -1;

    let beginMs = getStartDelayMs(timeNode);
    for (let index = 0; index < effectIndex; index++) beginMs += getStartDelayMs(ancestors[index]);

    let effect = NoTrackEffect;
    if (effectIndex >= 0) {
      const effectNode = ancestors[effectIndex];
      if (rawStage.trackEffects.has(effectNode)) {
        effect = rawStage.trackEffects.get(effectNode);
      } else {
        let effectBeginMs = 0;
        for (let index = effectIndex; index < ancestors.length; index++) effectBeginMs += getStartDelayMs(ancestors[index]);
        effect = {
          beginMs: effectBeginMs,
          durationMs: getEffectDurationMs(effectNode),
          repeatCount: getAttrValue(effectNode, 'repeatCount') ?? '',
          accel: getAccelDecelFraction(effectNode, 'accel'),
          decel: getAccelDecelFraction(effectNode, 'decel'),
        };
        rawStage.trackEffects.set(effectNode, effect);
      }
    }

    return {
      beginMs,
      durationMs: durationMsOf(timeNode),
      autoReverse: isTrueAttribute(getAttrValue(timeNode, 'autoRev')),
      holds: !eqIgnoreCase(getAttrValue(timeNode, 'fill'), 'remove'),
      accel: getAccelDecelFraction(timeNode, 'accel'),
      decel: getAccelDecelFraction(timeNode, 'decel'),
      timeFilter: parseTimeFilter(getAttrValue(timeNode, 'tmFilter')),
      effect,
    };
  },

  // A fade on the opacity track (emphasis fades, exit fades and exits without a clip model).
  _addFadeTrackSegments(rawStage, animationEffect, fadeOut) {
    this._addTrackSegments(rawStage, animationEffect, (timing) =>
      trackSegment({ property: 'opacity', kind: 'fade', fadeOut }, timing)
    );
  },

  _addTrackSegments(rawStage, behaviour, build) {
    const segment = build(this._getTrackTiming(rawStage, behaviour));
    for (const target of this._getTargetShapeIds(behaviour)) {
      if (!rawStage.trackSegmentsByTarget.has(target)) rawStage.trackSegmentsByTarget.set(target, []);
      rawStage.trackSegmentsByTarget.get(target).push(segment);
    }
  },

  // A two-keyframe ppt_w / ppt_h animation between multiples of the target's own size (Zoom,
  // Grow & Turn: 0 → #ppt_w) becomes a scale track, a style.rotation one between two angles a
  // rotate track. True when the animation is handled that way (or doesn't change anything).
  _addSizeOrRotationTrack(rawStage, animation, attributeName) {
    const width = eqIgnoreCase(attributeName, 'ppt_w');
    const height = eqIgnoreCase(attributeName, 'ppt_h');
    const rotation = eqIgnoreCase(attributeName, 'style.rotation');
    if (!width && !height && !rotation) return false;
    const keyframes = this._getPositionKeyframes(animation);
    if (keyframes == null) return false;

    const fromValue = parseLinearFormula(keyframes.from);
    const toValue = parseLinearFormula(keyframes.to);
    if (fromValue == null || toValue == null) return false;

    let fromFactor;
    let toFactor;
    if (rotation) {
      if (!formulaIsConstant(fromValue) || !formulaIsConstant(toValue)) return false;
      fromFactor = fromValue.constant;
      toFactor = toValue.constant;
    } else {
      fromFactor = tryGetSizeFactor(fromValue, width);
      toFactor = tryGetSizeFactor(toValue, width);
      if (fromFactor == null || toFactor == null) return false;
    }

    if (Math.abs(fromFactor - toFactor) < 1e-9) return true;

    this._addTrackSegments(rawStage, animation, (timing) =>
      trackSegment(
        {
          property: rotation ? 'rotate' : 'scale',
          kind: 'fromTo',
          axis: rotation || width ? 0 : 1,
          fromX: fromFactor,
          fromY: fromFactor,
          toX: toFactor,
          toY: toFactor,
        },
        timing
      )
    );
    return true;
  },

  // The keyframes of a linearly interpolated <p:anim> with their values and fmla, or null.
  _readFormulaKeyframes(animation) {
    const calcMode = getAttrValue(animation, 'calcmode') ?? 'lin';
    if (!eqIgnoreCase(calcMode, 'lin')) return null;

    const tavList = pFirstChild(animation, 'tavLst');
    const keyframes = tavList == null ? null : childElements(tavList).filter((el) => isP(el, 'tav'));
    if (keyframes == null || keyframes.length === 0) return null;

    const result = [];
    for (let index = 0; index < keyframes.length; index++) {
      const time = parseAnimationIndex(getAttrValue(keyframes[index], 'tm'));
      const value = getTimingValueText(keyframes[index]);
      if (value.length === 0) return null;
      result.push({
        time: time >= 0 ? time / 100000 : keyframes.length > 1 ? index / (keyframes.length - 1) : 0,
        value,
        formula: getAttrValue(keyframes[index], 'fmla') ?? '',
      });
    }
    return result;
  },

  // The stage's track segments by target, plus the position animations the Fly In model
  // couldn't take (as translate segments) when every one of them could be read.
  _collectTrackSegments(rawStage) {
    const result = new Map();
    for (const [target, segments] of rawStage.trackSegmentsByTarget) result.set(target, segments.slice());

    for (const [target, position] of rawStage.positionAnimationsByTarget) {
      if (position.isSupported || !position.tracksReadable || position.trackSegments.length === 0) continue;
      if (result.has(target)) result.get(target).push(...position.trackSegments);
      else result.set(target, position.trackSegments.slice());
    }
    return result;
  },

  _buildShapeAnimationCssInfo(shapeAnimationInfo, target = '', shapeBoxes = EmptyShapeBoxes) {
    const currentStageAnimationParts = [];
    const persistentAnimationParts = [];
    let style = '';

    const entranceDurationSeconds = this.toCssDurationSeconds(shapeAnimationInfo.entranceDurationMs, 0.5);
    const entranceDuration = fmtTrim(entranceDurationSeconds, 3);
    const widthAnimation = shapeAnimationInfo.numericAnimationsByAttributeName.get('ppt_w');
    const hasWidthAnimation = widthAnimation != null && widthAnimation.changesValue;
    // A repeating effect replays its fade with it (PowerPoint restarts every behaviour of the
    // effect), except for the swivels, whose pptxRotateYSwivelLoop keeps the shape on.
    const loopsEntrance =
      this._isIndefiniteRepeat(shapeAnimationInfo.entranceRepeatCount) &&
      !(hasWidthAnimation && widthAnimation.hasSineExpression);
    const loop = loopsEntrance ? ' infinite' : '';
    // This stage's part waits for its delay (After Previous…); a loop carries on in later stages
    // without one.
    const entranceTiming = shapeAnimationInfo.entranceTiming + formatAnimationDelay(shapeAnimationInfo.entranceBeginMs) + loop;
    const persistentTiming = shapeAnimationInfo.entranceTiming + loop;
    const entranceFilter = shapeAnimationInfo.entranceFilter;
    const wipeDirection = getWipeDirection(entranceFilter);
    const clipFilterStem = getClipFilterStem(entranceFilter);
    const barsOrientation = getRandomBarsOrientation(entranceFilter);
    let entranceName = '';
    let entranceFill = 'both';
    if (entranceFilter.toLowerCase().includes('fade')) {
      entranceName = 'pptxFadeIn';
    } else if (barsOrientation.length > 0) {
      currentStageAnimationParts.push(`pptxRandomBarsIn ${entranceDuration}s ${entranceTiming} both`);
      style = buildRandomBarsClipStyle(target, barsOrientation, shapeBoxes, (length, seedText) =>
        this.buildDeterministicShuffle(length, seedText)
      );
    } else if (wipeDirection.length > 0) {
      entranceName = 'pptxWipeIn' + wipeDirection;
    } else if (clipFilterStem.length > 0) {
      // Only before and while it plays: once revealed nothing (outline, shadow) is clipped.
      entranceName = 'pptxFilterIn' + clipFilterStem;
      entranceFill = 'backwards';
    } else if (eqIgnoreCase(shapeAnimationInfo.entranceTransition, 'in')) {
      entranceName = 'pptxFadeIn';
    }

    if (entranceName.length > 0) {
      currentStageAnimationParts.push(`${entranceName} ${entranceDuration}s ${entranceTiming} ${entranceFill}`);
      if (loopsEntrance) persistentAnimationParts.push(`${entranceName} ${entranceDuration}s ${persistentTiming} ${entranceFill}`);
    }

    // An exit's clip filter plays backwards and keeps the target concealed until it is hidden.
    const exitFilter = shapeAnimationInfo.exitFilter;
    if (exitFilter.length > 0) {
      const exitDuration = fmtTrim(this.toCssDurationSeconds(shapeAnimationInfo.exitDurationMs, 0.5), 3);
      const exitTiming = shapeAnimationInfo.exitTiming + formatAnimationDelay(shapeAnimationInfo.exitBeginMs);
      const exitBars = getRandomBarsOrientation(exitFilter);
      const exitWipe = getWipeDirection(exitFilter);
      const exitName =
        exitBars.length > 0 ? 'pptxRandomBarsOut' : exitWipe.length > 0 ? 'pptxWipeOut' + exitWipe : 'pptxFilterOut' + getClipFilterStem(exitFilter);
      currentStageAnimationParts.push(`${exitName} ${exitDuration}s ${exitTiming} both`);
      if (exitBars.length > 0) {
        style = buildRandomBarsClipStyle(target, exitBars, shapeBoxes, (length, seedText) =>
          this.buildDeterministicShuffle(length, seedText)
        );
      }
    }

    if (hasWidthAnimation && !widthAnimation.isTrack) {
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

    return { currentStageAnimationParts, persistentAnimationParts, style };
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
  // real stage when it carries an effect this converter actually renders: an entrance/exit (<p:set>
  // visibility toggle / <p:animEffect>), a numeric emphasis (<p:anim>), a motion path
  // (<p:animMotion>), a rotation (<p:animRot>), a scale (<p:animScale>) or a colour change
  // (<p:animClr>). Without one, the emitted
  // sub-slide would be a byte-for-byte duplicate of the main slide (e.g. audio-2-files was
  // producing slideN.html plus an identical slideN_1.html per audio clip).
  _stageHasRenderedEffect(stageElement) {
    return descendants(stageElement).some(
      (element) =>
        isP(element, 'set') ||
        isP(element, 'animEffect') ||
        isP(element, 'anim') ||
        isP(element, 'animMotion') ||
        isP(element, 'animRot') ||
        isP(element, 'animScale') ||
        isP(element, 'animClr')
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
