'use strict';

// Port of PptxToHtml/pptx/PptxConverter.cs (the Run entry point, slide HTML
// assembly, and random-bars transition helpers).

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { firstChild, elements, descendants, attr, childElements, innerText, NS } = require('../ooxml/dom');
const { notesSlidePart } = require('../ooxml/parts');
const { StringBuilder } = require('../util/stringbuilder');
const { f0, bankRound } = require('../util/format');
const { serializeIndented } = require('../util/dotnetJson');
const { placeholderOf, textBodyOf, shapeTreeOf } = require('./elements');
const {
  RandomBarsRevealDurationFraction,
  RandomBarsVerticalChunkTargetPx,
  RandomBarsHorizontalChunkTargetPx,
  RandomBarsChunkMinScale,
  RandomBarsChunkMaxScale,
} = require('./constants');

const MASK64 = (1n << 64n) - 1n;

function distinctOic(list) {
  const seen = new Set();
  const out = [];
  for (const s of list) {
    const k = s.toLowerCase();
    if (!seen.has(k)) {
      seen.add(k);
      out.push(s);
    }
  }
  return out;
}

module.exports = {
  run(presentationPart, slideRelIds, outputDir, inputPath, toolName, toolVersion, exportedAt, eot2ttfPath = null) {
    this._outputDir = outputDir;
    this._eot2ttfToolPath = eot2ttfPath;
    this._imageCounter = 0;
    this._mediaCounter = 0;
    this._savedImagePaths.clear();
    this._savedMediaPaths.clear();
    this._allFonts.clear();
    this._missingFonts.clear();
    this._embeddedFontFamilyNames.clear();
    fs.mkdirSync(path.join(outputDir, 'fonts'), { recursive: true });
    fs.mkdirSync(path.join(outputDir, 'images'), { recursive: true });
    fs.mkdirSync(path.join(outputDir, 'media'), { recursive: true });

    this.extractEmbeddedFonts(presentationPart);
    console.log(`  ${this._embeddedPptxFonts.size} embedded font(s) found.`);

    // Slide dimensions (default 10"x7.5" in EMU).
    let slideWidthEmu = 9144000;
    let slideHeightEmu = 6858000;
    const slideSize = firstChild(presentationPart.root, 'sldSz');
    if (slideSize != null) {
      if (attr(slideSize, 'cx') != null) slideWidthEmu = Number(attr(slideSize, 'cx'));
      if (attr(slideSize, 'cy') != null) slideHeightEmu = Number(attr(slideSize, 'cy'));
    }

    const slideWidthPx = this.emuToPx(slideWidthEmu);
    const slideHeightPx = this.emuToPx(slideHeightEmu);
    console.log(`Slide size: ${f0(slideWidthPx)}x${f0(slideHeightPx)}px`);
    console.log();

    const slideInfos = [];
    for (let i = 0; i < slideRelIds.length; i++) {
      process.stdout.write(`[${i + 1}/${slideRelIds.length}] Converting slide ${i + 1}... `);
      const previousSlidePart = i > 0 ? presentationPart.getPartById(slideRelIds[i - 1]) : null;
      const slidePart = presentationPart.getPartById(slideRelIds[i]);
      const animationInfo = this.extractSlideAnimationInfo(slidePart);
      const transitionInfo = this.extractSlideTransitionInfo(slidePart);
      const transitionRenderInfo = this.extractSlideTransitionShapeRenderInfo(previousSlidePart, slidePart, transitionInfo);

      this.loadThemeColors(slidePart);

      this._currentSlideImages.length = 0;
      this._currentSlideVideos.length = 0;
      this._currentSlideAudios.length = 0;

      const slideBaseName = `slide${i + 1}`;
      this._currentSlideDataIdPrefix = `${i + 1}`;
      this._currentSlideDataIdCounter = 0;
      const html = this.convertSlideToHtml(
        slidePart,
        presentationPart,
        i + 1,
        slideWidthPx,
        slideHeightPx,
        animationInfo.initiallyHiddenShapeIds,
        transitionRenderInfo.animatedShapeStylesByShapeId,
        transitionRenderInfo.pictureOverlaysByShapeId,
        previousSlidePart,
        transitionInfo
      );
      const outPath = path.join(outputDir, `${slideBaseName}.html`);
      fs.writeFileSync(outPath, '﻿' + html, 'utf8');

      const subHtmlFileNames = [];
      for (const stage of animationInfo.stages) {
        this._currentSlideDataIdPrefix = `${i + 1}_${stage.stageNumber}`;
        this._currentSlideDataIdCounter = 0;
        const subSlideHtml = this.convertSlideToHtml(
          slidePart,
          presentationPart,
          i + 1,
          slideWidthPx,
          slideHeightPx,
          stage.hiddenShapeIdsAfterStage,
          stage.animatedShapeStylesByShapeId
        );
        const subSlideHtmlFileName = `${slideBaseName}_${stage.stageNumber}.html`;
        fs.writeFileSync(path.join(outputDir, subSlideHtmlFileName), '﻿' + subSlideHtml, 'utf8');
        subHtmlFileNames.push(subSlideHtmlFileName);
      }

      const stageSummary = subHtmlFileNames.length === 0 ? '' : `, ${subHtmlFileNames.length} stage(s)`;
      console.log(`done -> ${path.basename(outPath)} (${Math.floor(html.length / 1024)} KB${stageSummary})`);

      slideInfos.push({
        htmlFileName: path.basename(outPath),
        subHtmlFileNames,
        transition: this.buildSlideTransitionJson(transitionInfo),
        isDisabled: this.isSlideDisabled(slidePart),
        images: distinctOic(this._currentSlideImages),
        videos: distinctOic(this._currentSlideVideos),
        audios: distinctOic(this._currentSlideAudios),
        note: this.getSlideNoteText(slidePart),
      });
    }

    console.log();

    if (this._missingFonts.count > 0) {
      console.error(`Missing fonts (${this._missingFonts.count}): ${this._missingFonts.join(', ')}`);
    }

    const fileBytes = fs.readFileSync(inputPath);
    const md5 = crypto.createHash('md5').update(fileBytes).digest('hex');
    const sha256 = crypto.createHash('sha256').update(fileBytes).digest('hex');

    const width = bankRound(slideWidthPx);
    const height = bankRound(slideHeightPx);
    const infoPath = path.join(outputDir, 'info.json');
    const pptxFileName = path.basename(inputPath);
    const info = {
      toolName,
      toolVersion,
      exportedAt,
      pptxFileName,
      dimensions: { width, height },
      checksum: { sha256, md5 },
      fontFamily: this._allFonts.toArray(),
      embeddedFontFamily: this._embeddedFontFamilyNames.toArray(),
      missingFontFamily: this._missingFonts.toArray(),
      slides: slideInfos,
    };
    const infoJson = serializeIndented(info);
    fs.writeFileSync(infoPath, '﻿' + infoJson + '\n', 'utf8');
    console.log(`Metadata written -> ${path.basename(infoPath)}`);

    console.log(`Done. ${slideRelIds.length} slide(s) converted to ${outputDir}`);
  },

  appendStageAnimationCss(sb) {
    sb.appendLine('@keyframes pptxFadeIn { from { opacity:0; } to { opacity:1; } }');
    sb.appendLine('@keyframes pptxSlideTranslateIn { from { transform:translate(var(--pptx-slide-translate-from-x, 0%), var(--pptx-slide-translate-from-y, 0%)); } to { transform:translate(0%, 0%); } }');
    sb.appendLine('@keyframes pptxClipInFromLeft { from { opacity:0.35; clip-path:inset(0 100% 0 0); } to { opacity:1; clip-path:inset(0 0 0 0); } }');
    sb.appendLine('@keyframes pptxClipInFromRight { from { opacity:0.35; clip-path:inset(0 0 0 100%); } to { opacity:1; clip-path:inset(0 0 0 0); } }');
    sb.appendLine('@keyframes pptxClipInFromTop { from { opacity:0.35; clip-path:inset(100% 0 0 0); } to { opacity:1; clip-path:inset(0 0 0 0); } }');
    sb.appendLine('@keyframes pptxClipInFromBottom { from { opacity:0.35; clip-path:inset(0 0 100% 0); } to { opacity:1; clip-path:inset(0 0 0 0); } }');
    sb.appendLine('@keyframes pptxSplitInVertical { from { opacity:0.4; clip-path:inset(0 50% 0 50%); } to { opacity:1; clip-path:inset(0 0 0 0); } }');
    sb.appendLine('@keyframes pptxSplitInHorizontal { from { opacity:0.4; clip-path:inset(50% 0 50% 0); } to { opacity:1; clip-path:inset(0 0 0 0); } }');
    sb.appendLine('@keyframes pptxCircleIn { from { opacity:0.3; clip-path:circle(0% at 50% 50%); } to { opacity:1; clip-path:circle(75% at 50% 50%); } }');
    sb.appendLine('@keyframes pptxRandomBarsReveal { from { opacity:0; } to { opacity:1; } }');
    sb.appendLine('@keyframes pptxMorphTransform { from { transform:translate(var(--pptx-morph-translate-x, 0px), var(--pptx-morph-translate-y, 0px)) rotate(var(--pptx-morph-rotate-from, 0deg)) scale(var(--pptx-morph-scale-x, 1), var(--pptx-morph-scale-y, 1)); } to { transform:translate(0px, 0px) rotate(0deg) scale(1, 1); } }');
    sb.appendLine('@keyframes pptxMorphFadeOut { from { opacity:1; } to { opacity:0; } }');
    sb.appendLine('@keyframes pptxHorizontalBarsIn { 0% { opacity:0; clip-path:inset(0 0 100% 0); } 100% { opacity:1; clip-path:inset(0 0 0 0); } }');
    sb.appendLine('@keyframes pptxRotateYSwivel { 0% { transform:rotateY(90deg); } 20% { transform:rotateY(0deg); } 40% { transform:rotateY(-90deg); } 60% { transform:rotateY(-180deg); } 80% { transform:rotateY(-270deg); } 100% { transform:rotateY(-360deg); } }');
    sb.appendLine('@keyframes pptxRotateYSwivelLoop { 0% { transform:rotateY(0deg); } 25% { transform:rotateY(90deg); } 50% { transform:rotateY(180deg); } 75% { transform:rotateY(270deg); } 100% { transform:rotateY(360deg); } }');
    sb.appendLine('@keyframes pptxScaleXSine { 0% { transform:scaleX(0); } 20% { transform:scaleX(1); } 40% { transform:scaleX(0); } 60% { transform:scaleX(-1); } 80% { transform:scaleX(0); } 100% { transform:scaleX(1); } }');
    sb.appendLine('@keyframes pptxScaleXSineLoop { 0% { transform:scaleX(1); } 25% { transform:scaleX(0); } 50% { transform:scaleX(-1); } 75% { transform:scaleX(0); } 100% { transform:scaleX(1); } }');
    sb.appendLine('@keyframes pptxScaleXGrow { from { transform:scaleX(0.75); } to { transform:scaleX(1); } }');
    sb.appendLine('@keyframes pptxScaleXGrowLoop { 0% { transform:scaleX(0.75); } 50% { transform:scaleX(1); } 100% { transform:scaleX(0.75); } }');
  },

  convertSlideToHtml(
    slidePart,
    presentationPart,
    slideNumber,
    slideWidthPx,
    slideHeightPx,
    hiddenShapeIds = null,
    animatedShapeStylesByShapeId = null,
    morphPictureOverlaysByShapeId = null,
    previousSlidePart = null,
    slideTransitionInfo = null
  ) {
    const slideTransitionStyle = this.buildSlideTransitionStyle(slideTransitionInfo);
    const renderRandomBars = slideTransitionInfo != null && (slideTransitionInfo.type || '').toLowerCase() === 'random';
    const renderPreviousSlideLayer =
      previousSlidePart != null && (slideTransitionStyle.trim() !== '' || renderRandomBars);
    const slideTransitionAttributes =
      slideTransitionInfo != null && slideTransitionInfo.type && slideTransitionInfo.type.trim() !== ''
        ? ` data-pptx-transition-type="${slideTransitionInfo.type}" data-pptx-transition-duration-ms="${Math.max(0, slideTransitionInfo.durationMs)}"`
        : '';

    const sb = new StringBuilder();
    sb.appendLine('<!DOCTYPE html>');
    sb.appendLine('<html lang="en">');
    sb.appendLine('<head>');
    sb.appendLine('<meta charset="UTF-8">');
    sb.appendLine(`<title>Slide ${slideNumber}</title>`);
    sb.appendLine('</head>');
    sb.appendLine(`<body style="width:${f0(slideWidthPx)}px;height:${f0(slideHeightPx)}px;overflow:hidden;margin:0;padding:0;top:0;left:0;">`);
    sb.appendLine('<style>');

    let usedFonts = this.collectFontsFromSlide(slidePart);
    if (renderPreviousSlideLayer && previousSlidePart != null) {
      usedFonts = usedFonts.concat(this.collectFontsFromSlide(previousSlidePart));
    }

    for (const fontName of distinctOic(usedFonts)) {
      this._allFonts.add(fontName);
      const fontFaceCss = this.getFontFaceCss(fontName);
      if (fontFaceCss != null) sb.appendLine(fontFaceCss);
    }

    sb.appendLine('* { margin: 0; padding: 0; box-sizing: border-box; }');
    sb.appendLine('table { border-collapse: collapse; width: 100%; height: 100%; }');
    sb.appendLine('td { border: 1px solid #999; padding: 4px; vertical-align: top; }');
    sb.appendLine(`.slide { position:relative; width:${f0(slideWidthPx)}px; height:${f0(slideHeightPx)}px; overflow:hidden; }`);
    sb.appendLine('.pptx-slide-layer { position:absolute; inset:0; display:block; overflow:hidden; }');
    if (
      (animatedShapeStylesByShapeId ? animatedShapeStylesByShapeId.size : 0) > 0 ||
      slideTransitionStyle.trim() !== '' ||
      renderRandomBars
    )
      this.appendStageAnimationCss(sb);
    sb.appendLine('</style>');

    const baseSvgStyle = 'display:block;overflow:hidden;position:absolute;inset:0;';
    const currentSvgStyle = baseSvgStyle + 'z-index:1;' + slideTransitionStyle;

    sb.appendLine(`<div class="slide"${slideTransitionAttributes}>`);
    if (renderPreviousSlideLayer && previousSlidePart != null) {
      const previousSvgStyle = baseSvgStyle + 'z-index:0;';
      sb.append(this.buildSlideSvgMarkup(previousSlidePart, presentationPart, slideWidthPx, slideHeightPx, previousSvgStyle));
    }

    if (renderRandomBars && slideTransitionInfo != null) {
      sb.append(
        this.buildRandomBarsSlideSvgMarkup(
          slidePart,
          presentationPart,
          slideWidthPx,
          slideHeightPx,
          hiddenShapeIds,
          animatedShapeStylesByShapeId,
          morphPictureOverlaysByShapeId,
          slideTransitionInfo
        )
      );
      this.appendRandomBarsScript(sb);
    } else {
      sb.append(
        this.buildSlideSvgMarkup(
          slidePart,
          presentationPart,
          slideWidthPx,
          slideHeightPx,
          currentSvgStyle,
          hiddenShapeIds,
          animatedShapeStylesByShapeId,
          morphPictureOverlaysByShapeId
        )
      );
    }
    sb.appendLine('</div>');
    sb.appendLine('</body>');
    sb.appendLine('</html>');
    return sb.toString();
  },

  buildSlideSvgMarkup(
    slidePart,
    presentationPart,
    slideWidthPx,
    slideHeightPx,
    svgStyle,
    hiddenShapeIds = null,
    animatedShapeStylesByShapeId = null,
    morphPictureOverlaysByShapeId = null,
    additionalSvgAttributes = ''
  ) {
    const originalHiddenShapeIds = this._currentHiddenShapeIds;
    const originalAnimatedShapeStylesByShapeId = this._currentAnimatedShapeStylesByShapeId;
    const originalMorphPictureOverlaysByShapeId = this._currentMorphPictureOverlaysByShapeId;

    this._currentHiddenShapeIds = hiddenShapeIds ? new Set(hiddenShapeIds) : new Set();
    this._currentAnimatedShapeStylesByShapeId = animatedShapeStylesByShapeId || new Map();
    this._currentMorphPictureOverlaysByShapeId = morphPictureOverlaysByShapeId || new Map();

    try {
      this._defsCounter = 0;
      const defsSb = new StringBuilder();
      const bodySb = new StringBuilder();

      const bgSvg = this.getSlideBackgroundSvg(slidePart, presentationPart, slideWidthPx, slideHeightPx, defsSb);
      bodySb.appendLine(bgSvg);

      const spTree = shapeTreeOf(slidePart.root);
      if (spTree != null) {
        this.renderShapeTree(bodySb, defsSb, spTree, slidePart);
      }

      const svgSb = new StringBuilder();
      svgSb.appendLine(
        `<svg class="pptx-slide-layer" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${f0(slideWidthPx)}" height="${f0(slideHeightPx)}" viewBox="0 0 ${f0(slideWidthPx)} ${f0(slideHeightPx)}" style="${svgStyle}"${additionalSvgAttributes}>`
      );
      if (defsSb.length > 0) {
        svgSb.appendLine('<defs>');
        svgSb.append(defsSb.toString());
        svgSb.appendLine('</defs>');
      }

      svgSb.append(bodySb.toString());
      svgSb.appendLine('</svg>');
      return svgSb.toString();
    } finally {
      this._currentHiddenShapeIds = originalHiddenShapeIds;
      this._currentAnimatedShapeStylesByShapeId = originalAnimatedShapeStylesByShapeId;
      this._currentMorphPictureOverlaysByShapeId = originalMorphPictureOverlaysByShapeId;
    }
  },

  buildRandomBarsSlideSvgMarkup(
    slidePart,
    presentationPart,
    slideWidthPx,
    slideHeightPx,
    hiddenShapeIds,
    animatedShapeStylesByShapeId,
    morphPictureOverlaysByShapeId,
    slideTransitionInfo
  ) {
    const isHorizontal = (slideTransitionInfo.orientation || '').toLowerCase() === 'horizontal';
    const randomBarSegments = this.buildRandomBarSegments(
      isHorizontal ? slideHeightPx : slideWidthPx,
      isHorizontal ? RandomBarsHorizontalChunkTargetPx : RandomBarsVerticalChunkTargetPx,
      isHorizontal ? 14 : 18,
      isHorizontal ? 28 : 30,
      `${slidePart.uri}|${slideTransitionInfo.orientation}|${f0(slideWidthPx)}x${f0(slideHeightPx)}|segments`
    );
    const randomBarOrder = this.buildDeterministicShuffle(
      randomBarSegments.length,
      `${slidePart.uri}|${slideTransitionInfo.orientation}|${f0(slideWidthPx)}x${f0(slideHeightPx)}|order`
    );
    const durationMs = Math.max(1, slideTransitionInfo.durationMs);
    const revealDurationMs = Math.max(1, bankRound(durationMs * RandomBarsRevealDurationFraction));
    const barStepDurationMs = Math.max(1, Math.trunc(revealDurationMs / Math.max(1, randomBarOrder.length)));
    const sb = new StringBuilder();

    for (let revealPosition = 0; revealPosition < randomBarOrder.length; revealPosition++) {
      const barIndex = randomBarOrder[revealPosition];
      const delayMs = barStepDurationMs * revealPosition;

      const [x, y, width, height] = this.getRandomBarRect(randomBarSegments[barIndex], isHorizontal, slideWidthPx, slideHeightPx);

      const barSvgStyle =
        'display:block;overflow:hidden;position:absolute;inset:0;z-index:1;opacity:0;will-change:opacity,clip-path;' +
        this.buildRandomBarClipStyle(x, y, width, height, slideWidthPx, slideHeightPx);

      sb.append(
        this.buildSlideSvgMarkup(
          slidePart,
          presentationPart,
          slideWidthPx,
          slideHeightPx,
          barSvgStyle,
          hiddenShapeIds,
          animatedShapeStylesByShapeId,
          morphPictureOverlaysByShapeId,
          ` data-pptx-random-bar-layer="1" data-pptx-random-delay-ms="${delayMs}"`
        )
      );
    }

    return sb.toString();
  },

  appendRandomBarsScript(sb) {
    sb.appendLine('<script>');
    sb.appendLine('(() => {');
    sb.appendLine('  const showRandomBars = () => {');
    sb.appendLine('    const randomBarLayers = document.querySelectorAll(\'[data-pptx-random-bar-layer="1"]\');');
    sb.appendLine('    for (const layer of randomBarLayers) {');
    sb.appendLine("      const delayMs = Number.parseInt(layer.getAttribute('data-pptx-random-delay-ms') || '0', 10);");
    sb.appendLine('      window.setTimeout(() => {');
    sb.appendLine("        layer.style.opacity = '1';");
    sb.appendLine('      }, Number.isFinite(delayMs) ? Math.max(0, delayMs) : 0);');
    sb.appendLine('    }');
    sb.appendLine('  };');
    sb.appendLine("  if (document.readyState === 'loading') {");
    sb.appendLine("    document.addEventListener('DOMContentLoaded', showRandomBars, { once: true });");
    sb.appendLine('    return;');
    sb.appendLine('  }');
    sb.appendLine('  showRandomBars();');
    sb.appendLine('})();');
    sb.appendLine('</script>');
  },

  getRandomBarRect(segment, isHorizontal, slideWidthPx, slideHeightPx) {
    if (isHorizontal) return [0.0, segment.startPx, slideWidthPx, segment.lengthPx];
    return [segment.startPx, 0.0, segment.lengthPx, slideHeightPx];
  },

  buildRandomBarSegments(totalPx, targetPx, minCount, maxCount, seedText) {
    const segmentCount = Math.max(minCount, Math.min(maxCount, Math.round(totalPx / targetPx)));
    if (segmentCount <= 0) return [];

    let state = { value: createDeterministicSeed(seedText) };
    const weights = new Array(segmentCount);
    let totalWeight = 0.0;
    for (let i = 0; i < segmentCount; i++) {
      const weight = RandomBarsChunkMinScale + (RandomBarsChunkMaxScale - RandomBarsChunkMinScale) * nextDeterministicDouble(state);
      weights[i] = weight;
      totalWeight += weight;
    }

    const segments = new Array(segmentCount);
    let startPx = 0.0;
    let remainingPx = totalPx;
    let remainingWeight = totalWeight;
    for (let i = 0; i < segmentCount; i++) {
      const lengthPx = i === segmentCount - 1 ? remainingPx : Math.max(1.0, remainingPx * (weights[i] / remainingWeight));
      segments[i] = { startPx, lengthPx };
      startPx += lengthPx;
      remainingPx -= lengthPx;
      remainingWeight -= weights[i];
    }

    return segments;
  },

  buildRandomBarClipStyle(x, y, width, height, slideWidthPx, slideHeightPx) {
    const top = y;
    const right = Math.max(0.0, slideWidthPx - (x + width));
    const bottom = Math.max(0.0, slideHeightPx - (y + height));
    const left = x;
    const inset = `inset(${formatSvgNumber(top)}px ${formatSvgNumber(right)}px ${formatSvgNumber(bottom)}px ${formatSvgNumber(left)}px)`;
    return `clip-path:${inset};-webkit-clip-path:${inset};`;
  },

  buildDeterministicShuffle(length, seedText) {
    const values = Array.from({ length }, (_, i) => i);
    if (length <= 1) return values;

    const state = { value: createDeterministicSeed(seedText) };

    for (let i = values.length - 1; i > 0; i--) {
      state.value = nextDeterministicState(state.value);
      const swapIndex = Number(state.value % BigInt(i + 1));
      const tmp = values[i];
      values[i] = values[swapIndex];
      values[swapIndex] = tmp;
    }

    return values;
  },

  getSlideNoteText(slidePart) {
    const notesPart = notesSlidePart(slidePart);
    const shapeTree = notesPart ? shapeTreeOf(notesPart.root) : null;
    if (shapeTree == null) return null;

    const textBlocks = [];
    for (const shape of elements(shapeTree, 'sp', NS.P)) {
      const placeholderType = attr(placeholderOf(shape), 'type');

      if (placeholderType === 'sldNum' || placeholderType === 'dt' || placeholderType === 'ftr' || placeholderType === 'hdr') {
        continue;
      }

      const txBody = textBodyOf(shape);
      const shapeText = descendants(txBody, 't', NS.A)
        .map((t) => innerText(t))
        .filter((t) => t != null && t.trim() !== '')
        .join('\n')
        .trim();

      if (shapeText.trim() !== '') textBlocks.push(shapeText);
    }

    if (textBlocks.length === 0) return null;

    return textBlocks.join('\n');
  },

  isSlideDisabled(slidePart) {
    const slide = slidePart.root;
    if (slide == null) return false;

    for (let i = 0; i < slide.attributes.length; i++) {
      const a = slide.attributes[i];
      if (a.localName.toLowerCase() !== 'show') continue;
      const value = a.value || '';
      return value === '0' || value.toLowerCase() === 'false';
    }

    return false;
  },

  renderShapeTree(sb, defsSb, shapeTree, slidePart) {
    for (const child of childElements(shapeTree)) {
      if (child.localName === 'sp') this.renderShape(sb, defsSb, child, slidePart);
      else if (child.localName === 'pic') this.renderPicture(sb, defsSb, child, slidePart);
      else if (child.localName === 'graphicFrame') this.renderGraphicFrame(sb, defsSb, child, slidePart);
      else if (child.localName === 'grpSp') this.renderGroupShape(sb, defsSb, child, slidePart);
      else if (child.localName === 'cxnSp') this.renderConnectionShape(sb, defsSb, child, slidePart);
    }
  },
};

// ── deterministic RNG (parity with C# CreateDeterministicSeed etc.) ─────────

function readUInt64LE(buf, offset) {
  let v = 0n;
  for (let i = 0; i < 8; i++) v |= BigInt(buf[offset + i]) << BigInt(8 * i);
  return v;
}

function createDeterministicSeed(seedText) {
  const seedBytes = crypto.createHash('sha256').update(Buffer.from(seedText, 'utf8')).digest();
  const state = (readUInt64LE(seedBytes, 0) ^ readUInt64LE(seedBytes, 8)) & MASK64;
  return state === 0n ? 0x9e3779b97f4a7c15n : state;
}

function nextDeterministicState(state) {
  return (state * 6364136223846793005n + 1442695040888963407n) & MASK64;
}

function nextDeterministicDouble(stateRef) {
  stateRef.value = nextDeterministicState(stateRef.value);
  return Number(stateRef.value & 0xffffffffn) / 4294967295;
}

function formatSvgNumber(value) {
  // C# value.ToString("0.###", InvariantCulture)
  const { fmtTrim } = require('../util/format');
  return fmtTrim(value, 3);
}
