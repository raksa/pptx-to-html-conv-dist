'use strict';

// Numeric constants shared across converter modules (from PptxConverter.cs /
// TextRenderer.cs).

const pkg = require('../../package.json');

module.exports = {
  // CSS line-height ratio fallbacks (see PptxConverter.cs comments).
  DefaultPptLineHeightRatio: 1.2,
  DefaultLineHeightRatio: 1.4,
  // Share of the CSS half-leading by which the first line of a top-anchored text box
  // with expanded line spacing is pushed down (see TextRenderer).
  FirstLineLeadingShiftFactor: 0.5,
  // Baseline of a WordArt-warped line, as a share of its band from the top (see
  // TextRenderer.renderWarpedTextBody).
  WarpBaselineRatio: 0.8,

  // Random-bars slide transition tuning.
  RandomBarsRevealDurationFraction: 0.55,
  RandomBarsVerticalChunkTargetPx: 56.0,
  RandomBarsHorizontalChunkTargetPx: 40.0,
  RandomBarsChunkMinScale: 0.65,
  RandomBarsChunkMaxScale: 1.35,

  // TextRenderer shadow calibration.
  TextShadowDistanceScale: 0.78,
  TextShadowBlurScale: 0.72,
  TextShadowOpacityScale: 0.68,

  // `pkg.name` is the npm id ("pptx-to-html-conv"); info.json must carry the same
  // tool name the C# port writes, so it stays a literal here.
  ToolName: 'PptxToHtml',
  ToolVersion: pkg.version,
};
