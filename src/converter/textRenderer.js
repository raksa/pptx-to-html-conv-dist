'use strict';

// Port of PptxToHtml/pptx/TextRenderer.cs

const { firstChild, elements, attr, childElements, innerText, NS } = require('../ooxml/dom');
const { StringBuilder } = require('../util/stringbuilder');
const { f0, f1, f2, f3 } = require('../util/format');
const { PRESET_COLORS } = require('./colorHelper');
const {
  DefaultPptLineHeightRatio,
  HalfLeadingCompensationFactor,
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

    for (const para of elements(textBody, 'p', A)) {
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
          if (bulletColor === '') bulletColor = fb.color;
          if (bulletFont === '') bulletFont = fb.font;
          if (bulletSizePct === 1.0 && fb.sizePct !== 1.0) bulletSizePct = fb.sizePct;
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

      // Line spacing.
      let lnSpcStyle = '';
      let cssLineHeightRatio = 0;
      const paraLnSpc = firstChild(pPr, 'lnSpc');
      const inhLnSpc = firstChild(effectiveBulletSource, 'lnSpc');
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
      const effectiveSpcBef = firstChild(pPr, 'spcBef') || firstChild(effectiveBulletSource, 'spcBef');
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
          if (spcBefPt > 0) spacingStyle = `margin-top:${f1(spcBefPt)}pt;`;
        } else if (spcPts != null && attr(spcPts, 'val') != null) {
          const pts = num(spcPts, 'val') / 100.0;
          if (pts > 0) spacingStyle = `margin-top:${f1(pts)}pt;`;
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

      if (isFirstParagraph && compensateLeading && cssLineHeightRatio > paraLHRatio) {
        let maxFontPt = 0;
        for (const fpRun of elements(para, 'r', A)) {
          const fpRp = firstChild(fpRun, 'rPr');
          if (fpRp != null && attr(fpRp, 'sz') != null) maxFontPt = Math.max(maxFontPt, (num(fpRp, 'sz') / 100.0) * fontScale);
        }
        if (maxFontPt === 0 && effectiveDefRunProps != null && attr(effectiveDefRunProps, 'sz') != null)
          maxFontPt = (num(effectiveDefRunProps, 'sz') / 100.0) * fontScale;

        if (maxFontPt > 0) {
          const halfLeadingPt = ((cssLineHeightRatio - paraLHRatio) / 2.0) * maxFontPt * HalfLeadingCompensationFactor;
          combinedParaStyle += (combinedParaStyle !== '' ? ';' : '') + `margin-top:-${f1(halfLeadingPt)}pt`;
        }
      }

      // Per-segment (break-separated) max run font sizes.
      const segmentMaxFontPts = [];
      {
        let sMax = 0;
        for (const child of childElements(para)) {
          if (child.localName === 'r' && attr(firstChild(child, 'rPr'), 'sz') != null) {
            sMax = Math.max(sMax, (num(firstChild(child, 'rPr'), 'sz') / 100.0) * fontScale);
          } else if (child.localName === 'br') {
            segmentMaxFontPts.push(sMax);
            sMax = 0;
          }
        }
        segmentMaxFontPts.push(sMax);
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

      const makeParaAttr = (baseStyle, segIdx) => {
        const segLnSpc = getSegmentLnSpc(segIdx);
        let s = baseStyle;
        if (segLnSpc !== '') s += (s !== '' ? ';' : '') + segLnSpc;
        return s !== '' ? ` style="${s}"` : '';
      };

      let currentSegment = 0;
      const paraTextShadow = this.getCssPropertyValue(combinedParaStyle, 'text-shadow');
      sb.append(`<p${makeParaAttr(combinedParaStyle, 0)}>`);

      const runsAny = elements(para, 'r', A).length > 0;

      if (hasBullet && runsAny) {
        let buStyle = '';
        if (bulletFont !== '') buStyle += `font-family:${this.buildCssFontFamily(bulletFont)};`;
        if (bulletColor !== '') buStyle += `color:${bulletColor};`;
        if (bulletSizePct !== 1.0) buStyle += `font-size:${f0(bulletSizePct * 100)}%;`;

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
            const text = this.encode(innerText(firstChild(child, 't')) || '');
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
            const text = this.encode(innerText(firstChild(child, 't')) || '');

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
            sb.appendLine('</p>');
            currentSegment++;
            sb.append(`<p${makeParaAttr(combinedParaStyle, currentSegment)}>`);
          }
        }
        this._flushPendingRun(sb, pendingStyle, pendingText);
      }

      const breaksAny = elements(para, 'br', A).length > 0;
      if (!runsAny && !breaksAny) {
        sb.append('&nbsp;');
      }

      sb.appendLine('</p>');
      isFirstParagraph = false;
    }
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
    }

    return styles.join(';');
  },

  _getShadowCssColor(outerShdw) {
    let hexWithOptionalAlpha = '00000080';

    const srgb = firstChild(outerShdw, 'srgbClr');
    if (srgb != null && attr(srgb, 'val') != null) {
      hexWithOptionalAlpha = this.applyColorTransforms(attr(srgb, 'val'), childElements(srgb));
    } else {
      const scheme = firstChild(outerShdw, 'schemeClr');
      if (scheme != null && attr(scheme, 'val') != null) {
        const hex = this.resolveSchemeColor(attr(scheme, 'val'));
        hexWithOptionalAlpha = this.applyColorTransforms(hex, childElements(scheme));
      } else {
        const preset = firstChild(outerShdw, 'prstClr');
        if (preset != null && attr(preset, 'val') != null && PRESET_COLORS.has(attr(preset, 'val'))) {
          const mapped = PRESET_COLORS.get(attr(preset, 'val'));
          const rgb = mapped.startsWith('#') ? mapped.slice(1) : mapped;
          hexWithOptionalAlpha = this.applyColorTransforms(rgb, childElements(preset));
        }
      }
    }

    return this._cssColorFromHexWithAlpha(hexWithOptionalAlpha);
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
