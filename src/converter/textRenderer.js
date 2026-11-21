'use strict';

// Port of PptxToHtml/pptx/TextRenderer.cs

const { firstChild, elements, attr, childElements, innerText, NS } = require('../ooxml/dom');
const { StringBuilder } = require('../util/stringbuilder');
const { f0, f1, f2, f3, fmtTrim } = require('../util/format');
const { PRESET_COLORS } = require('./colorHelper');
const {
  DefaultPptLineHeightRatio,
  FirstLineLeadingShiftFactor,
  WarpBaselineRatio,
  TextShadowDistanceScale,
  TextShadowBlurScale,
  TextShadowOpacityScale,
} = require('./constants');

const A = NS.A;

function num(el, name, fallback = null) {
  const v = attr(el, name);
  if (v == null) return fallback;
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

// HTML collapses runs of these characters (white-space: normal), so a line whose
// text is made only of them renders with no height at all.
function isCollapsibleWhitespace(text) {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c !== 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0c && c !== 0x0d) return false;
  }
  return true;
}

// Fill choices of text run properties (EG_FillProperties), which the schema places after <a:ln>.
const TEXT_FILL_ELEMENTS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

// Numeric attribute of an effect element, or the fallback when absent or not a number
// (mirrors ReadEffectNumber's double.TryParse in the C# tool).
function readEffectNumber(el, name, fallback) {
  const v = attr(el, name);
  if (v == null || v.trim() === '') return fallback;
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

module.exports = {
  renderTextBody(sb, textBody, inherited, fontScale = 1.0, lnSpcReduction = 0.0, compensateLeading = false, noWrap = false, isVerticalText = false) {
    const localLstStyle = firstChild(textBody, 'lstStyle');
    let isFirstParagraph = true;
    const autoNumCounters = new Map();
    const autoNumTypes = new Map();

    // First explicit alignment used as the text-body default.
    let textBodyDefaultAlign = '';
    for (const p of elements(textBody, 'p', A)) {
      const a = this.getAlignment(firstChild(p, 'pPr'));
      if (a !== '') { textBodyDefaultAlign = a; break; }
    }

    let paragraphIndex = -1;
    for (const para of elements(textBody, 'p', A)) {
      paragraphIndex++;
      const pPr = firstChild(para, 'pPr');
      const paraLevel = num(pPr, 'lvl', 0) || 0;
      const levelProps = this.getLevelParagraphProperties(localLstStyle, paraLevel);
      const inheritedLevelProps = this.getLevelPropsFromComposite(inherited.bulletStyleSource, paraLevel);
      const effectiveBulletSource = levelProps || inheritedLevelProps || inherited.bulletParaProps;
      const effectiveDefRunProps =
        this.getUsefulDefRunProps(levelProps) || this.getUsefulDefRunProps(inheritedLevelProps) || inherited.runProps;

      let align = this.getAlignment(pPr);

      if (align === '' && levelProps != null && attr(levelProps, 'algn') != null) align = this.getAlignmentToken(attr(levelProps, 'algn'));
      if (align === '' && inheritedLevelProps != null && attr(inheritedLevelProps, 'algn') != null)
        align = this.getAlignmentToken(attr(inheritedLevelProps, 'algn'));
      if (align === '' && inherited.alignment != null) align = this.getAlignmentToken(inherited.alignment);
      if (align === '' && textBodyDefaultAlign !== '') align = textBodyDefaultAlign;

      // Bullet resolution.
      let hasBullet = false;
      let bulletChar = '';
      let bulletColor = '';
      let bulletFont = '';
      let bulletSizePct = 1.0;
      const noBulletExplicit = firstChild(pPr, 'buNone') != null;

      if (!noBulletExplicit) {
        const pBuChar = firstChild(pPr, 'buChar');
        const pBuAutoNum = firstChild(pPr, 'buAutoNum');

        if (pBuChar != null) {
          hasBullet = true;
          bulletChar = this.getStringValue(attr(pBuChar, 'char'), '■');
          const bs = this._resolveBulletStyling(pPr, bulletColor, bulletSizePct, bulletFont);
          bulletColor = bs.color; bulletSizePct = bs.sizePct; bulletFont = bs.font;
          const fb = this._resolveBulletStyling(effectiveBulletSource, '', 1.0, '');
          // buClrTx / buFontTx / buSzTx on the paragraph mean "follow the text", which
          // overrides whatever the inherited list style would supply.
          if (bulletColor === '' && firstChild(pPr, 'buClrTx') == null) bulletColor = fb.color;
          if (bulletFont === '' && firstChild(pPr, 'buFontTx') == null) bulletFont = fb.font;
          if (bulletSizePct === 1.0 && fb.sizePct !== 1.0 && firstChild(pPr, 'buSzTx') == null) bulletSizePct = fb.sizePct;
        } else if (pBuAutoNum != null) {
          hasBullet = true;
          const curType = attr(pBuAutoNum, 'type');
          if (!autoNumTypes.has(paraLevel) || autoNumTypes.get(paraLevel) !== curType) {
            autoNumCounters.set(paraLevel, 0);
            autoNumTypes.set(paraLevel, curType);
          }
          autoNumCounters.set(paraLevel, autoNumCounters.get(paraLevel) + 1);
          const startAt = num(pBuAutoNum, 'startAt', 1);
          const numValue = startAt + autoNumCounters.get(paraLevel) - 1;
          bulletChar = this._formatAutoNumber(curType, numValue);
          if (pPr != null) {
            const buClr = firstChild(pPr, 'buClr');
            if (buClr != null) {
              const buSchemeClr = firstChild(buClr, 'schemeClr');
              if (buSchemeClr != null && attr(buSchemeClr, 'val') != null) {
                let bHex = this.resolveSchemeColor(attr(buSchemeClr, 'val'));
                bHex = this.applyColorTransforms(bHex, childElements(buSchemeClr));
                bulletColor = `#${bHex}`;
              }
            }
            const buSzPct = firstChild(pPr, 'buSzPct');
            if (buSzPct != null && attr(buSzPct, 'val') != null) bulletSizePct = num(buSzPct, 'val') / 100000.0;
            const buFont = firstChild(pPr, 'buFont');
            bulletFont = this._resolveBulletFontTypeface(buFont);
          }
        } else if (effectiveBulletSource != null) {
          const buAutoNum = firstChild(effectiveBulletSource, 'buAutoNum');
          const buChar = firstChild(effectiveBulletSource, 'buChar');
          if (buAutoNum != null) {
            hasBullet = true;
            const curType = attr(buAutoNum, 'type');
            if (!autoNumTypes.has(paraLevel) || autoNumTypes.get(paraLevel) !== curType) {
              autoNumCounters.set(paraLevel, 0);
              autoNumTypes.set(paraLevel, curType);
            }
            autoNumCounters.set(paraLevel, autoNumCounters.get(paraLevel) + 1);
            const startAt = num(buAutoNum, 'startAt', 1);
            const numValue = startAt + autoNumCounters.get(paraLevel) - 1;
            bulletChar = this._formatAutoNumber(curType, numValue);
            const bs = this._resolveBulletStyling(effectiveBulletSource, bulletColor, bulletSizePct, bulletFont);
            bulletColor = bs.color; bulletSizePct = bs.sizePct; bulletFont = bs.font;
          } else if (buChar != null) {
            hasBullet = true;
            bulletChar = this.getStringValue(attr(buChar, 'char'), '■');
            const bs = this._resolveBulletStyling(effectiveBulletSource, bulletColor, bulletSizePct, bulletFont);
            bulletColor = bs.color; bulletSizePct = bs.sizePct; bulletFont = bs.font;
          }
        }
      }

      // Merged default run style across the cascade.
      const fallbackLevelProps = this.getLevelPropsFromComposite(inherited.fallbackStyleSource, paraLevel);
      const defaultStyle = this.getMergedRunStyles(
        fontScale,
        firstChild(levelProps, 'defRPr'),
        firstChild(inheritedLevelProps, 'defRPr'),
        firstChild(fallbackLevelProps, 'defRPr'),
        inherited.runProps
      );

      // Paragraph spacing is inherited one property at a time: a list-style level that exists
      // but omits lnSpc/spcBef must not hide the value a lower level (e.g. the master's body
      // style under a layout placeholder's own level) defines.
      const spacingSources = [levelProps, inheritedLevelProps, fallbackLevelProps, inherited.bulletParaProps];

      // Line spacing.
      let lnSpcStyle = '';
      let cssLineHeightRatio = 0;
      const paraLnSpc = firstChild(pPr, 'lnSpc');
      const inhLnSpc = this._findInheritedParagraphChild(spacingSources, 'lnSpc');
      const effectiveLnSpc = paraLnSpc || inhLnSpc;

      const paraLHRatio = this.getParagraphLineHeightRatio(para, effectiveDefRunProps);

      if (effectiveLnSpc != null) {
        const spcPct = firstChild(effectiveLnSpc, 'spcPct');
        const spcPts = firstChild(effectiveLnSpc, 'spcPts');
        if (spcPct != null && attr(spcPct, 'val') != null) {
          const pct = Math.max(0.0, num(spcPct, 'val') / 100000.0 - lnSpcReduction);
          if (noWrap && Math.abs(pct - 1.0) < 0.0001) {
            cssLineHeightRatio = 1.0;
            lnSpcStyle = 'line-height:1.00;';
          } else {
            const cssLH = pct * DefaultPptLineHeightRatio;
            cssLineHeightRatio = cssLH;
            lnSpcStyle = `line-height:${f2(cssLH)};`;
          }
        } else if (spcPts != null && attr(spcPts, 'val') != null) {
          const pts = (num(spcPts, 'val') / 100.0) * (1.0 - lnSpcReduction);
          lnSpcStyle = `line-height:${f1(pts)}pt;`;
        }
      } else {
        const cssLH = Math.max(0.0, 1.0 - lnSpcReduction) * DefaultPptLineHeightRatio;
        cssLineHeightRatio = cssLH;
        lnSpcStyle = `line-height:${f2(cssLH)};`;
      }

      // Space before.
      let spacingStyle = '';
      // The space before in pt, also on a first paragraph (where it isn't laid out): a
      // paragraph's animation box includes it.
      let spaceBeforePt = 0;
      const effectiveSpcBef = firstChild(pPr, 'spcBef') || this._findInheritedParagraphChild(spacingSources, 'spcBef');
      if (effectiveSpcBef != null) {
        const spcPct = firstChild(effectiveSpcBef, 'spcPct');
        const spcPts = firstChild(effectiveSpcBef, 'spcPts');
        if (spcPct != null && attr(spcPct, 'val') != null) {
          const spcBefPct = num(spcPct, 'val') / 100000.0;
          let spcMaxFontPt = 0;
          for (const spcRun of elements(para, 'r', A)) {
            const rp = firstChild(spcRun, 'rPr');
            if (rp != null && attr(rp, 'sz') != null) spcMaxFontPt = Math.max(spcMaxFontPt, (num(rp, 'sz') / 100.0) * fontScale);
          }
          if (spcMaxFontPt === 0 && effectiveDefRunProps != null && attr(effectiveDefRunProps, 'sz') != null)
            spcMaxFontPt = (num(effectiveDefRunProps, 'sz') / 100.0) * fontScale;
          const spcBefPt = spcBefPct * DefaultPptLineHeightRatio * spcMaxFontPt;
          if (spcBefPt > 0) {
            spacingStyle = `margin-top:${f1(spcBefPt)}pt;`;
            spaceBeforePt = spcBefPt;
          }
        } else if (spcPts != null && attr(spcPts, 'val') != null) {
          const pts = num(spcPts, 'val') / 100.0;
          if (pts > 0) {
            spacingStyle = `margin-top:${f1(pts)}pt;`;
            spaceBeforePt = pts;
          }
        }
      }

      // Left margin / indent.
      let bulletMargin = '';
      const isCentered = align.includes('center');
      if (!isCentered && !isVerticalText) {
        let marL = 0;
        let indent = 0;
        const otherLevelProps = this.getLevelPropsFromComposite(this._masterOtherStyle, paraLevel);
        if (attr(pPr, 'marL') != null) marL = parseInt(attr(pPr, 'marL'), 10);
        else if (attr(effectiveBulletSource, 'marL') != null) marL = parseInt(attr(effectiveBulletSource, 'marL'), 10);
        else if (attr(fallbackLevelProps, 'marL') != null) marL = parseInt(attr(fallbackLevelProps, 'marL'), 10);
        else if (attr(otherLevelProps, 'marL') != null) marL = parseInt(attr(otherLevelProps, 'marL'), 10);

        if (hasBullet) {
          if (attr(pPr, 'indent') != null) indent = parseInt(attr(pPr, 'indent'), 10);
          else if (attr(effectiveBulletSource, 'indent') != null) indent = parseInt(attr(effectiveBulletSource, 'indent'), 10);
          else if (attr(fallbackLevelProps, 'indent') != null) indent = parseInt(attr(fallbackLevelProps, 'indent'), 10);
          else if (attr(otherLevelProps, 'indent') != null) indent = parseInt(attr(otherLevelProps, 'indent'), 10);
          const marLPx = this.emuToPx(marL);
          const indentPx = this.emuToPx(Math.abs(indent));
          bulletMargin = `margin-left:${f0(marLPx)}px;text-indent:-${f0(indentPx)}px;`;
        } else if (marL !== 0) {
          bulletMargin = `margin-left:${f0(this.emuToPx(marL))}px;`;
        }
      }

      let combinedParaStyle = align;
      if (defaultStyle !== '') combinedParaStyle += (combinedParaStyle !== '' ? ';' : '') + defaultStyle;
      if (bulletMargin !== '') combinedParaStyle += (combinedParaStyle !== '' ? ';' : '') + bulletMargin;
      if (!isFirstParagraph && spacingStyle !== '') combinedParaStyle += (combinedParaStyle !== '' ? ';' : '') + spacingStyle;

      // First line of a top-anchored text box with expanded line spacing. CSS centres each
      // line in its line box (half the extra leading above it, half below); PowerPoint puts
      // more of it above the first line. Measured against PowerPoint's own video export
      // (130–300% spacing), its first line sits a further half of the CSS half-leading lower
      // while later lines keep the CSS pitch, so only the first paragraph is shifted down.
      if (isFirstParagraph && compensateLeading && cssLineHeightRatio > paraLHRatio) {
        let maxFontPt = 0;
        for (const fpRun of elements(para, 'r', A)) {
          const fpRp = firstChild(fpRun, 'rPr');
          if (fpRp != null && attr(fpRp, 'sz') != null) maxFontPt = Math.max(maxFontPt, (num(fpRp, 'sz') / 100.0) * fontScale);
        }
        if (maxFontPt === 0 && effectiveDefRunProps != null && attr(effectiveDefRunProps, 'sz') != null)
          maxFontPt = (num(effectiveDefRunProps, 'sz') / 100.0) * fontScale;

        if (maxFontPt > 0) {
          const halfLeadingPt = ((cssLineHeightRatio - paraLHRatio) / 2.0) * maxFontPt;
          combinedParaStyle += (combinedParaStyle !== '' ? ';' : '') + `margin-top:${f1(halfLeadingPt * FirstLineLeadingShiftFactor)}pt`;
        }
      }

      // Per-segment (break-separated) metrics. PowerPoint gives every line a height even when
      // it shows no text: such a line is sized by its runs, or failing that by the <a:br> that
      // ends it (the paragraph's last line: by <a:endParaRPr>). HTML collapses a line with no
      // visible text to nothing, so those segments get an explicit &nbsp; filler.
      const segmentMaxFontPts = [];
      const segmentFillerFontPts = [];
      const segmentIsBlank = [];
      {
        let sMax = 0;
        let sBlank = true;
        const closeSegment = (terminatorRPr) => {
          const terminatorPt = attr(terminatorRPr, 'sz') != null ? (num(terminatorRPr, 'sz') / 100.0) * fontScale : 0;
          segmentMaxFontPts.push(sBlank && sMax === 0 ? terminatorPt : sMax);
          segmentFillerFontPts.push(sMax > 0 ? sMax : terminatorPt);
          segmentIsBlank.push(sBlank);
          sMax = 0;
          sBlank = true;
        };
        for (const child of childElements(para)) {
          if (child.localName === 'r') {
            if (attr(firstChild(child, 'rPr'), 'sz') != null)
              sMax = Math.max(sMax, (num(firstChild(child, 'rPr'), 'sz') / 100.0) * fontScale);
            if (!isCollapsibleWhitespace(innerText(firstChild(child, 't')) || '')) sBlank = false;
          } else if (child.localName === 'br') {
            closeSegment(firstChild(child, 'rPr'));
          }
        }
        closeSegment(firstChild(para, 'endParaRPr'));
      }

      let pDefFontPt = 0;
      {
        const dfz = this.getCssPropertyValue(combinedParaStyle, 'font-size');
        if (dfz.endsWith('pt')) {
          const parsed = parseFloat(dfz.substring(0, dfz.length - 2));
          if (!Number.isNaN(parsed)) pDefFontPt = parsed;
        }
      }

      const getSegmentLnSpc = (segIdx) => {
        if (cssLineHeightRatio > 0 && segIdx < segmentMaxFontPts.length) {
          const segMax = segmentMaxFontPts[segIdx];
          if (segMax > 0 && pDefFontPt > 0 && segMax < pDefFontPt - 0.5) {
            const fixedLH = segMax * cssLineHeightRatio;
            return `line-height:${f1(fixedLH)}pt;`;
          }
        }
        return lnSpcStyle;
      };

      // Paragraph-level animation state (a "build by paragraph" effect): not yet revealed,
      // entering in this stage, or moved by a motion path.
      const paragraphAnimationStyle = this._buildParagraphAnimationStyle(paragraphIndex, spaceBeforePt);

      const makeParaAttr = (baseStyle, segIdx) => {
        const segLnSpc = getSegmentLnSpc(segIdx);
        let s = baseStyle;
        if (segLnSpc !== '') s += (s !== '' ? ';' : '') + segLnSpc;
        if (paragraphAnimationStyle !== '') s += (s !== '' && !s.endsWith(';') ? ';' : '') + paragraphAnimationStyle;
        return s !== '' ? ` style="${s}"` : '';
      };

      let currentSegment = 0;
      const paraTextShadow = this.getCssPropertyValue(combinedParaStyle, 'text-shadow');
      sb.append(`<p${makeParaAttr(combinedParaStyle, 0)}>`);

      const runs = elements(para, 'r', A);
      const runsAny = runs.length > 0;
      const bulletShown = hasBullet && runsAny;

      const appendBlankSegmentFiller = (segIdx) => {
        // The bullet already gives the first line of a bulleted paragraph its height.
        if (!segmentIsBlank[segIdx] || (bulletShown && segIdx === 0)) return;
        // A line no larger than the paragraph default is already sized by the <p> itself
        // (getSegmentLnSpc pins smaller lines to a pt line-height); a smaller span here
        // would only add baseline-alignment slack against the <p> strut. Larger lines
        // need a span of their own size to grow the line box.
        const fillerPt = segmentFillerFontPts[segIdx];
        if (fillerPt > 0 && (pDefFontPt <= 0 || fillerPt > pDefFontPt + 0.5))
          sb.append(`<span style="font-size:${f1(fillerPt)}pt">&nbsp;</span>`);
        else sb.append('&nbsp;');
      };

      if (bulletShown) {
        // Bullet color, size and font follow the paragraph's first run unless the bullet sets
        // its own (PowerPoint's buClrTx / buSzTx / buFontTx defaults); buSzPct scales that size.
        const firstRunProps = firstChild(runs[0], 'rPr');
        const firstRunStyle = this.getRunStyle(firstRunProps, fontScale);
        let buStyle = '';
        const buFontFamily = bulletFont !== '' ? this.buildCssFontFamily(bulletFont) : this.getCssPropertyValue(firstRunStyle, 'font-family');
        if (buFontFamily !== '') buStyle += `font-family:${buFontFamily};`;
        const buColor = bulletColor !== '' ? bulletColor : this.getCssPropertyValue(firstRunStyle, 'color');
        if (buColor !== '' && buColor !== 'transparent') buStyle += `color:${buColor};`;
        if (attr(firstRunProps, 'sz') != null) buStyle += `font-size:${f1((num(firstRunProps, 'sz') / 100.0) * fontScale * bulletSizePct)}pt;`;
        else if (bulletSizePct !== 1.0) buStyle += `font-size:${f0(bulletSizePct * 100)}%;`;

        const bulletText = this.encode(bulletChar);
        if (buStyle !== '') sb.append(`<span style="${buStyle}">${bulletText}</span>`);
        else sb.append(`<span>${bulletText}</span>`);
        sb.append('&nbsp;');

        let pendingStyle = null;
        const pendingText = new StringBuilder();
        for (const child of childElements(para)) {
          if (child.localName === 'r') {
            let spanStyle = this.getRunStyle(firstChild(child, 'rPr'), fontScale);
            spanStyle = this.promoteShadowToDropShadowForHighlight(spanStyle, paraTextShadow);
            const rawText = innerText(firstChild(child, 't')) || '';
            const text = this._wrapTextReflection(firstChild(child, 'rPr'), this.encode(rawText), rawText);
            if (spanStyle === pendingStyle) {
              pendingText.append(text);
            } else {
              this._flushPendingRun(sb, pendingStyle, pendingText);
              pendingStyle = spanStyle;
              pendingText.append(text);
            }
          } else if (child.localName === 'br') {
            this._flushPendingRun(sb, pendingStyle, pendingText);
            pendingStyle = null;
            appendBlankSegmentFiller(currentSegment);
            sb.appendLine('</p>');
            currentSegment++;
            sb.append(`<p${makeParaAttr(combinedParaStyle, currentSegment)}>`);
          }
        }
        this._flushPendingRun(sb, pendingStyle, pendingText);
      } else {
        let pendingStyle = null;
        const pendingText = new StringBuilder();

        for (const child of childElements(para)) {
          if (child.localName === 'r') {
            const rPr = firstChild(child, 'rPr');
            let spanStyle = this.getRunStyle(rPr, fontScale);
            spanStyle = this.promoteShadowToDropShadowForHighlight(spanStyle, paraTextShadow);
            const rawText = innerText(firstChild(child, 't')) || '';
            const text = this._wrapTextReflection(firstChild(child, 'rPr'), this.encode(rawText), rawText);

            if (spanStyle === pendingStyle) {
              pendingText.append(text);
            } else {
              this._flushPendingRun(sb, pendingStyle, pendingText);
              pendingStyle = spanStyle;
              pendingText.append(text);
            }
          } else if (child.localName === 'br') {
            this._flushPendingRun(sb, pendingStyle, pendingText);
            pendingStyle = null;
            appendBlankSegmentFiller(currentSegment);
            sb.appendLine('</p>');
            currentSegment++;
            sb.append(`<p${makeParaAttr(combinedParaStyle, currentSegment)}>`);
          }
        }
        this._flushPendingRun(sb, pendingStyle, pendingText);
      }

      // An empty or whitespace-only last line still occupies a line
      appendBlankSegmentFiller(currentSegment);

      sb.appendLine('</p>');
      isFirstParagraph = false;
    }
  },

  /**
   * Style for paragraph `paragraphIndex` of the shape being rendered, from its
   * "<shape id>#<paragraph>" animation key. A paragraph not yet built in stays in the layout
   * but is invisible, as in PowerPoint, so revealing it never reflows the other lines.
   * A paragraph whose Fly In box the browser measures also gets its space before (px), which
   * PowerPoint counts in that box even on a first paragraph, where no margin shows it.
   */
  _buildParagraphAnimationStyle(paragraphIndex, spaceBeforePt = 0) {
    if (!this._currentTextShapeId) return '';

    const paragraphKey = `${this._currentTextShapeId}#${paragraphIndex}`;
    let style = this._currentHiddenShapeIds.has(paragraphKey) ? 'visibility:hidden;' : '';
    if (this._currentAnimatedShapeStylesByShapeId.has(paragraphKey)) {
      const animationStyle = this._currentAnimatedShapeStylesByShapeId.get(paragraphKey);
      if (spaceBeforePt > 0 && animationStyle.includes('--pptx-fly-box:'))
        style += `--pptx-space-before:${f1((spaceBeforePt * 96.0) / 72.0)}px;`;
      style += animationStyle;
    }
    return style;
  },

  /** First `childName` child found on the given paragraph-property sources, in order. */
  _findInheritedParagraphChild(sources, childName) {
    for (const source of sources) {
      const child = firstChild(source, childName);
      if (child != null) return child;
    }
    return null;
  },

  _flushPendingRun(sb, style, text) {
    if (text.length === 0) return;
    if (style != null && style !== '') sb.append(`<span style="${style}">${text.toString()}</span>`);
    else sb.append(text.toString());
    text.clear();
  },

  getRunStyle(rPr, fontScale = 1.0) {
    if (rPr == null) return '';

    const styles = [];

    const isHyperlinkRun = childElements(rPr).some((e) => {
      const ln = (e.localName || '').toLowerCase();
      return ln === 'hlinkclick' || ln === 'hlinkmouseover' || ln === 'hlink';
    });

    if (attr(rPr, 'sz') != null) {
      const ptSize = (num(rPr, 'sz') / 100.0) * fontScale;
      styles.push(`font-size:${f1(ptSize)}pt`);
    }

    if (attr(rPr, 'b') === '1' || (attr(rPr, 'b') || '').toLowerCase() === 'true') styles.push('font-weight:bold');
    if (attr(rPr, 'i') === '1' || (attr(rPr, 'i') || '').toLowerCase() === 'true') styles.push('font-style:italic');

    const textDecorations = [];
    const u = attr(rPr, 'u');
    if (u != null && u !== 'none') textDecorations.push('underline');
    const strike = attr(rPr, 'strike');
    if (strike != null && strike !== 'noStrike') textDecorations.push('line-through');
    if (textDecorations.length > 0) styles.push(`text-decoration:${textDecorations.join(' ')}`);

    if (attr(rPr, 'spc') != null) {
      const spacingPt = num(rPr, 'spc') / 100.0;
      styles.push(`letter-spacing:${f1(spacingPt)}pt`);
    }

    const solidFill = firstChild(rPr, 'solidFill');
    if (solidFill != null) {
      const color = this.getColorFromFill(solidFill);
      if (color !== '' && !isHyperlinkRun) styles.push(`color:${color}`);
    } else {
      const pattFill = firstChild(rPr, 'pattFill');
      if (pattFill != null) {
        const pattCss = this.getPatternFillCss(pattFill);
        if (pattCss !== '') {
          styles.push(pattCss.replace(/;+$/, ''));
          styles.push('-webkit-background-clip:text');
          styles.push('background-clip:text');
          styles.push('-webkit-text-fill-color:transparent');
          styles.push('color:transparent');
        }
      } else {
        const gradFill = firstChild(rPr, 'gradFill');
        if (gradFill != null) {
          const gradCss = this.getGradientFillCss(gradFill);
          if (gradCss !== '') {
            styles.push(gradCss.replace(/;+$/, ''));
            styles.push('-webkit-background-clip:text');
            styles.push('background-clip:text');
            styles.push('-webkit-text-fill-color:transparent');
            styles.push('color:transparent');
          }
        }
      }
    }

    if (isHyperlinkRun) {
      const linkHex = this.resolveSchemeColor('hlink');
      if (linkHex && linkHex.trim() !== '') styles.push(`color:#${linkHex}`);
    }

    const outlineStyle = this._getTextOutlineCss(rPr);
    if (outlineStyle !== '') styles.push(outlineStyle);

    const highlight = firstChild(rPr, 'highlight');
    if (highlight != null) {
      const hlColor = this.getColorFromHighlight(highlight);
      if (hlColor !== '') styles.push(`background-color:${hlColor}`);
    }

    const latin = firstChild(rPr, 'latin');
    const cs = firstChild(rPr, 'cs');
    const latinTypeface = this.getStringValue(attr(latin, 'typeface'));
    const csTypeface = this.getStringValue(attr(cs, 'typeface'));
    let typeface = latinTypeface !== '' ? latinTypeface : csTypeface;
    if (typeface === '+mj-lt') typeface = this._themeMajorFont;
    else if (typeface === '+mn-lt') typeface = this._themeMinorFont;
    else if (typeface === '+mj-cs' || typeface === '+mn-cs')
      typeface = csTypeface !== '' && csTypeface !== '+mn-cs' && csTypeface !== '+mj-cs' ? csTypeface : this._themeMinorFont;

    if (typeface === '' && this._themeMinorFont !== '') typeface = this._themeMinorFont;

    if (typeface !== '' && typeface !== '+mj-lt' && typeface !== '+mn-lt' && typeface !== '+mj-cs' && typeface !== '+mn-cs') {
      if (
        latinTypeface !== '' &&
        csTypeface !== '' &&
        latinTypeface !== csTypeface &&
        csTypeface !== '+mn-cs' &&
        csTypeface !== '+mj-cs'
      )
        styles.push(`font-family:${this.buildCssFontFamily(csTypeface, latinTypeface)}`);
      else styles.push(`font-family:${this.buildCssFontFamily(typeface)}`);
    }

    const effects = firstChild(rPr, 'effectLst');
    if (effects != null) {
      const outerShdw = firstChild(effects, 'outerShdw');
      if (outerShdw != null) {
        const blur = ((num(outerShdw, 'blurRad', 0) || 0) / 914400.0) * 96.0 * TextShadowBlurScale;
        const dist = ((num(outerShdw, 'dist', 0) || 0) / 914400.0) * 96.0 * TextShadowDistanceScale;
        const angle = (((num(outerShdw, 'dir', 0) || 0) / 60000.0) * Math.PI) / 180.0;
        const dx = dist * Math.cos(angle);
        const dy = dist * Math.sin(angle);
        const shadowColor = this._getShadowCssColor(outerShdw);
        styles.push(`text-shadow:${f1(dx)}px ${f1(dy)}px ${f1(blur)}px ${shadowColor}`);
      } else {
        styles.push('text-shadow:none');
      }

      const glowFilterId = this._getTextGlowFilterId(firstChild(effects, 'glow'));
      if (glowFilterId !== '') styles.push(`filter:url(#${glowFilterId})`);
    }

    return styles.join(';');
  },

  _getShadowCssColor(outerShdw) {
    // PowerPoint default shadow color is black at 50% opacity when not specified.
    return this._cssColorFromHexWithAlpha(this._resolveEffectColorHex(outerShdw, '00000080'));
  },

  /**
   * Color of a shadow/glow effect as RRGGBB plus an alpha byte when not opaque, or the
   * fallback when the effect names no color.
   */
  _resolveEffectColorHex(effect, fallbackHexWithAlpha) {
    const srgb = firstChild(effect, 'srgbClr');
    if (srgb != null && attr(srgb, 'val') != null) return this.applyColorTransforms(attr(srgb, 'val'), childElements(srgb));

    const scheme = firstChild(effect, 'schemeClr');
    if (scheme != null && attr(scheme, 'val') != null)
      return this.applyColorTransforms(this.resolveSchemeColor(attr(scheme, 'val')), childElements(scheme));

    const preset = firstChild(effect, 'prstClr');
    if (preset != null && attr(preset, 'val') != null && PRESET_COLORS.has(attr(preset, 'val'))) {
      const mapped = PRESET_COLORS.get(attr(preset, 'val'));
      const rgb = mapped.startsWith('#') ? mapped.slice(1) : mapped;
      return this.applyColorTransforms(rgb, childElements(preset));
    }

    return fallbackHexWithAlpha;
  },

  /**
   * Text outline (<a:ln> in run properties). PowerPoint strokes the glyph edge over the fill;
   * a CSS stroke painted over the fill also cancels Chrome's synthetic bold and hollows the
   * glyphs, so the stroke is painted behind the fill at twice the width instead, leaving the
   * same visible outline outside the glyph. w omitted means PowerPoint's 0.75pt default and
   * w="0" a hairline; an outline with no fill (or <a:noFill/>) draws nothing. The schema puts
   * <a:ln> before the text fill, and PowerPoint ignores an outline written after it.
   */
  _getTextOutlineCss(rPr) {
    let outline = null;
    for (const child of childElements(rPr)) {
      if (child.localName === 'ln') {
        outline = child;
        break;
      }
      if (TEXT_FILL_ELEMENTS.has(child.localName)) return '';
    }

    const outlineFill = firstChild(outline, 'solidFill');
    if (outlineFill == null) return '';

    const color = this.getColorFromFill(outlineFill);
    if (color === '') return '';

    const widthPx = Math.max((readEffectNumber(outline, 'w', 9525) / 914400.0) * 96.0, 0.5);
    return `-webkit-text-stroke:${f2(widthPx * 2)}px ${color};paint-order:stroke fill`;
  },

  /**
   * Glow (<a:glow rad=".."> + color): PowerPoint grows the glyph shape by the radius, softens
   * it and tints it. CSS text-shadow cannot grow a shape, so the glow is an SVG filter in the
   * current slide layer's <defs> (dilate by half the radius, blur by a quarter, flood with the
   * color), shared by every run with the same glow. Returns '' outside a slide layer.
   */
  _getTextGlowFilterId(glow) {
    if (glow == null || this._currentDefsSb == null) return '';

    const radiusPx = (readEffectNumber(glow, 'rad', 0) / 914400.0) * 96.0;
    if (radiusPx <= 0) return '';

    const hexWithAlpha = this._resolveEffectColorHex(glow, '000000');
    const colorHex = hexWithAlpha.length >= 6 ? hexWithAlpha.substring(0, 6) : '000000';
    const opacity = hexWithAlpha.length >= 8 ? parseInt(hexWithAlpha.substring(6, 8), 16) / 255.0 : 1.0;
    const spread = f2(radiusPx / 2);
    const blur = f2(radiusPx / 4);
    const opacityText = fmtTrim(opacity, 3);
    const key = `${spread}|${blur}|${colorHex}|${opacityText}`;
    if (this._textGlowFilterIds.has(key)) return this._textGlowFilterIds.get(key);

    const id = `glow${++this._defsCounter}`;
    this._currentDefsSb.appendLine(
      `<filter id="${id}" x="-50%" y="-100%" width="200%" height="300%" color-interpolation-filters="sRGB">` +
        `<feMorphology in="SourceAlpha" operator="dilate" radius="${spread}" result="spread"/>` +
        `<feGaussianBlur in="spread" stdDeviation="${blur}" result="soft"/>` +
        `<feFlood flood-color="#${colorHex}" flood-opacity="${opacityText}"/>` +
        '<feComposite in2="soft" operator="in" result="glow"/>' +
        '<feMerge><feMergeNode in="glow"/><feMergeNode in="SourceGraphic"/></feMerge></filter>'
    );
    this._textGlowFilterIds.set(key, id);
    return id;
  },

  /**
   * Reflection (<a:reflection>): the run is mirrored below itself, fading from stA to endA
   * between stPos and endPos (percent of its height) and offset by dist. The text goes in an
   * inline-block whose box ends at the font's descent line (line-height:normal), so
   * -webkit-box-reflect mirrors about the edge PowerPoint uses; the run's own span keeps the
   * paragraph line height, so the line layout does not change.
   */
  _wrapTextReflection(rPr, encodedText, rawText) {
    const reflection = firstChild(firstChild(rPr, 'effectLst'), 'reflection');
    if (reflection == null || isCollapsibleWhitespace(rawText)) return encodedText;

    const distancePx = (readEffectNumber(reflection, 'dist', 0) / 914400.0) * 96.0;
    const startAlpha = readEffectNumber(reflection, 'stA', 100000) / 100000.0;
    const endAlpha = readEffectNumber(reflection, 'endA', 0) / 100000.0;
    const startPos = readEffectNumber(reflection, 'stPos', 0) / 1000.0;
    const endPos = readEffectNumber(reflection, 'endPos', 100000) / 1000.0;

    // dir is the direction from the text to its reflection (clockwise from +x; 90° = below).
    const angle = (((readEffectNumber(reflection, 'dir', 5400000) / 60000.0) % 360) + 360) % 360;
    const [side, towards] =
      angle >= 45 && angle < 135
        ? ['below', 'bottom']
        : angle >= 135 && angle < 225
        ? ['left', 'left']
        : angle >= 225 && angle < 315
        ? ['above', 'top']
        : ['right', 'right'];

    // The mask is laid over the run before it is mirrored, so the end nearest the mirror
    // axis carries the start alpha.
    const farStop = fmtTrim(100 - endPos, 3);
    const nearStop = fmtTrim(100 - startPos, 3);
    const mask =
      `linear-gradient(to ${towards}, transparent ${farStop}%, ` +
      `rgba(0,0,0,${fmtTrim(endAlpha, 3)}) ${farStop}%, ` +
      `rgba(0,0,0,${fmtTrim(startAlpha, 3)}) ${nearStop}%)`;
    return `<span style="display:inline-block;line-height:normal;-webkit-box-reflect:${side} ${f1(distancePx)}px ${mask}">${encodedText}</span>`;
  },

  /**
   * Preset WordArt transform of a text body (<a:prstTxWarp prst="..">), or '' when the text
   * is not warped ("textNoShape" is the explicit no-warp preset).
   */
  getTextWarpPreset(bodyPr) {
    const warp = bodyPr == null ? null : childElements(bodyPr).find((child) => child.localName === 'prstTxWarp');
    const preset = attr(warp, 'prst') || '';
    return preset === 'textNoShape' ? '' : preset;
  },

  /**
   * WordArt transform (<a:prstTxWarp>): PowerPoint bends the text into the warp's envelope,
   * and the envelope fills the text box whatever the font size. Each line (paragraph or
   * <a:br> segment) is drawn as SVG text stretched to the box width (textLength) and sized to
   * an equal share of the box height; the envelope's own curve is not reproduced.
   */
  renderWarpedTextBody(sb, textBody, inherited, fontScale, width, height, leftInsetPx, topInsetPx, rightInsetPx, bottomInsetPx, defaultTextColor = null) {
    const textColor = defaultTextColor ?? this.getThemeTextColor();
    const localLstStyle = firstChild(textBody, 'lstStyle');
    const lines = [];
    for (const para of elements(textBody, 'p', A)) {
      const paraLevel = num(firstChild(para, 'pPr'), 'lvl', 0) || 0;
      const levelDefaults = firstChild(this.getLevelParagraphProperties(localLstStyle, paraLevel), 'defRPr');
      const inheritedDefaults = firstChild(this.getLevelPropsFromComposite(inherited.bulletStyleSource, paraLevel), 'defRPr');
      const fallbackDefaults = firstChild(this.getLevelPropsFromComposite(inherited.fallbackStyleSource, paraLevel), 'defRPr');
      let line = [];
      for (const child of childElements(para)) {
        if (child.localName === 'r') {
          const text = innerText(firstChild(child, 't')) || '';
          if (text === '') continue;
          const style = this.getMergedRunStyles(
            fontScale,
            firstChild(child, 'rPr'),
            levelDefaults,
            inheritedDefaults,
            fallbackDefaults,
            inherited.runProps
          );
          line.push({ style, text });
        } else if (child.localName === 'br') {
          lines.push(line);
          line = [];
        }
      }
      lines.push(line);
    }

    const innerWidth = width - leftInsetPx - rightInsetPx;
    const bandHeight = (height - topInsetPx - bottomInsetPx) / lines.length;
    if (innerWidth <= 0 || bandHeight <= 0) return;

    // The line's height (DefaultPptLineHeightRatio em) fills its band.
    const fontPx = bandHeight / DefaultPptLineHeightRatio;
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].every((part) => isCollapsibleWhitespace(part.text))) continue;

      const baseline = topInsetPx + bandHeight * i + bandHeight * WarpBaselineRatio;
      sb.append(
        `<text x="${f1(leftInsetPx)}" y="${f1(baseline)}" textLength="${f1(innerWidth)}" lengthAdjust="spacingAndGlyphs" style="font-size:${f1(fontPx)}px;fill:${textColor}">`
      );
      for (const part of lines[i]) sb.append(`<tspan style="${this._toSvgTextStyle(part.style)}">${this.encode(part.text)}</tspan>`);
      sb.appendLine('</text>');
    }
  },

  /**
   * Run CSS (as built for HTML text) translated for an SVG tspan: color becomes fill, the
   * text stroke becomes stroke/stroke-width, and text-shadow joins the filter list as a
   * drop-shadow (SVG text ignores text-shadow). Sizes and spacing are left to the warp.
   */
  _toSvgTextStyle(cssStyle) {
    const declarations = [];
    const filters = [];
    for (const declaration of cssStyle.split(';')) {
      const colon = declaration.indexOf(':');
      if (colon <= 0) continue;

      const name = declaration.substring(0, colon).trim().toLowerCase();
      const value = declaration.substring(colon + 1).trim();
      switch (name) {
        case 'color':
          declarations.push(`fill:${value}`);
          break;
        case 'font-family':
        case 'font-weight':
        case 'font-style':
        case 'text-decoration':
        case 'paint-order':
          declarations.push(`${name}:${value}`);
          break;
        case '-webkit-text-stroke': {
          const space = value.indexOf(' ');
          if (space > 0) {
            declarations.push(`stroke:${value.substring(space + 1)}`);
            declarations.push(`stroke-width:${value.substring(0, space)}`);
            declarations.push('stroke-linejoin:round');
          }
          break;
        }
        case 'text-shadow':
          if (value !== 'none') filters.push(`drop-shadow(${value})`);
          break;
        case 'filter':
          filters.push(value);
          break;
        default:
          break;
      }
    }

    if (filters.length > 0) declarations.push(`filter:${filters.join(' ')}`);
    return declarations.join(';');
  },

  _cssColorFromHexWithAlpha(hex) {
    if (hex == null || hex.trim() === '') return 'rgba(0,0,0,0.5)';
    const h = hex.trim().replace(/^#/, '');
    if (h.length < 6) return 'rgba(0,0,0,0.5)';

    const r = parseInt(h.substr(0, 2), 16);
    const g = parseInt(h.substr(2, 2), 16);
    const b = parseInt(h.substr(4, 2), 16);
    let a = 1.0;
    if (h.length >= 8) a = parseInt(h.substr(6, 2), 16) / 255.0;
    a = Math.max(0.0, Math.min(1.0, a * TextShadowOpacityScale));
    return `rgba(${r},${g},${b},${f3(a)})`;
  },

  promoteShadowToDropShadowForHighlight(style, paraShadow) {
    if (style === '' || !this.hasCssProperty(style, 'background-color')) return style;

    let shadow = this.getCssPropertyValue(style, 'text-shadow');
    if (shadow === '') shadow = paraShadow;
    if (shadow === '' || shadow.toLowerCase() === 'none') return style;

    if (this.hasCssProperty(style, 'filter')) return style;

    return style + `;filter:drop-shadow(${shadow})`;
  },

  hasCssProperty(style, propertyName) {
    return this.getCssPropertyValue(style, propertyName) !== '';
  },

  getCssPropertyValue(style, propertyName) {
    if (style === '') return '';
    const parts = style.split(';').filter((s) => s.length > 0);
    for (const raw of parts) {
      const colon = raw.indexOf(':');
      if (colon <= 0) continue;
      const key = raw.slice(0, colon).trim();
      if (key.toLowerCase() === propertyName.toLowerCase()) return raw.slice(colon + 1).trim();
    }
    return '';
  },

  getMergedRunStyles(fontScale, ...sources) {
    const props = new Map(); // lowercased property name -> full declaration
    for (let i = sources.length - 1; i >= 0; i--) {
      if (sources[i] == null) continue;
      const css = this.getRunStyle(sources[i], fontScale);
      for (const decl of css.split(';').filter((s) => s.length > 0)) {
        const colon = decl.indexOf(':');
        if (colon > 0) props.set(decl.slice(0, colon).trim().toLowerCase(), decl.trim());
      }
    }
    return Array.from(props.values()).join(';');
  },

  getAlignment(pPr) {
    const a = attr(pPr, 'algn');
    if (a == null) return '';
    return this.getAlignmentToken(a);
  },

  getAlignmentToken(alignVal) {
    if (alignVal === 'ctr') return 'text-align:center';
    if (alignVal === 'r') return 'text-align:right';
    if (alignVal === 'just') return 'text-align:justify';
    return '';
  },

  _resolveBulletStyling(source, bulletColor, bulletSizePct, bulletFont) {
    const result = { color: bulletColor, sizePct: bulletSizePct, font: bulletFont };
    if (source == null) return result;
    const buClr = firstChild(source, 'buClr');
    if (buClr != null) {
      const buSchemeClr = firstChild(buClr, 'schemeClr');
      if (buSchemeClr != null && attr(buSchemeClr, 'val') != null) {
        let bHex = this.resolveSchemeColor(attr(buSchemeClr, 'val'));
        bHex = this.applyColorTransforms(bHex, childElements(buSchemeClr));
        result.color = `#${bHex}`;
      }
      const buRgbClr = firstChild(buClr, 'srgbClr');
      if (buRgbClr != null && attr(buRgbClr, 'val') != null && result.color === '') result.color = `#${attr(buRgbClr, 'val')}`;
    }
    const buSzPct = firstChild(source, 'buSzPct');
    if (buSzPct != null && attr(buSzPct, 'val') != null) result.sizePct = num(buSzPct, 'val') / 100000.0;
    const buFont = firstChild(source, 'buFont');
    result.font = this._resolveBulletFontTypeface(buFont);
    return result;
  },

  _resolveBulletFontTypeface(buFont) {
    let typeface = this.getStringValue(attr(buFont, 'typeface'));
    if (typeface === '+mj-lt') typeface = this._themeMajorFont;
    else if (typeface === '+mn-lt') typeface = this._themeMinorFont;
    else if (typeface === '+mj-cs' || typeface === '+mn-cs') typeface = this._themeMinorFont;
    return typeface;
  },

  _formatAutoNumber(type, n) {
    if (type == null) return `${n}.`;
    switch (type) {
      case 'arabicPeriod': return `${n}.`;
      case 'arabicParenR': return `${n})`;
      case 'arabicParenBoth': return `(${n})`;
      case 'arabicPlain': return `${n}`;
      case 'romanUcPeriod': return `${toRoman(n)}.`;
      case 'romanLcPeriod': return `${toRoman(n).toLowerCase()}.`;
      case 'romanUcParenR': return `${toRoman(n)})`;
      case 'romanLcParenR': return `${toRoman(n).toLowerCase()})`;
      case 'romanUcParenBoth': return `(${toRoman(n)})`;
      case 'romanLcParenBoth': return `(${toRoman(n).toLowerCase()})`;
      case 'alphaUcPeriod': return `${toAlpha(n)}.`;
      case 'alphaLcPeriod': return `${toAlpha(n).toLowerCase()}.`;
      case 'alphaUcParenR': return `${toAlpha(n)})`;
      case 'alphaLcParenR': return `${toAlpha(n).toLowerCase()})`;
      case 'alphaUcParenBoth': return `(${toAlpha(n)})`;
      case 'alphaLcParenBoth': return `(${toAlpha(n).toLowerCase()})`;
      default: return `${n}.`;
    }
  },
};

function toRoman(num) {
  if (num <= 0 || num > 3999) return String(num);
  let out = '';
  const map = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ];
  for (const [value, symbol] of map) {
    while (num >= value) { out += symbol; num -= value; }
  }
  return out;
}

function toAlpha(num) {
  if (num <= 0) return String(num);
  let out = '';
  while (num > 0) {
    num--;
    out = String.fromCharCode(65 + (num % 26)) + out;
    num = Math.floor(num / 26);
  }
  return out;
}
