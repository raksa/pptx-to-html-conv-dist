'use strict';

// Port of the C# `static partial class PptxConverter`. The C# version keeps
// per-run state in static fields; here each conversion uses one instance whose
// prototype is composed from the ported method modules (mirroring the partial
// class split across files).

const { CIMap, CISet } = require('../util/collections');

class PptxConverter {
  constructor() {
    // Cache of font family name -> (winAscent+winDescent)/unitsPerEm ratio
    this._fontMetricsCache = new CIMap();
    // Cache of font family name -> ppt single-line height ratio
    this._fontPptLineHeightCache = new CIMap();

    // Theme color map loaded from the actual PPTX theme (scheme slot -> hex)
    this._themeColors = new Map();
    // Color map from slide master (scheme name -> theme color slot)
    this._colorMap = new Map();
    // Theme fonts
    this._themeMajorFont = '';
    this._themeMinorFont = '';

    // Master OtherStyle element (default paragraph-level props for text boxes).
    this._masterOtherStyle = null;
    // Theme line style widths from lnStyleLst (0-based).
    this._themeLineWidths = [];

    // Embedded font data extracted from the PPTX (font name -> Buffer)
    this._embeddedPptxFonts = new CIMap();
    // Cache of @font-face CSS rules (font name -> css string or null)
    this._fontFaceCache = new CIMap();
    // Fonts that could not be found anywhere
    this._missingFonts = new CISet();
    // All fonts referenced across all slides
    this._allFonts = new CISet();
    // Family names from the presentation EmbeddedFontList
    this._embeddedFontFamilyNames = new CISet();

    // Output directory for the current run
    this._outputDir = '';
    // eot2ttf tool path (or null)
    this._eot2ttfToolPath = null;
    // Unique filename counters
    this._imageCounter = 0;
    this._mediaCounter = 0;
    // Deduplication caches: part uri -> saved relative path
    this._savedImagePaths = new Map();
    this._savedMediaPaths = new Map();
    // Per-slide media tracking
    this._currentSlideImages = [];
    this._currentSlideVideos = [];
    this._currentSlideAudios = [];

    // Per-render animation/transition context (reset around each layer render).
    this._currentHiddenShapeIds = new Set();
    this._currentAnimatedShapeStylesByShapeId = new Map();
    this._currentMorphPictureOverlaysByShapeId = new Map();

    // SVG <defs> counter (unique ids within a single slide layer).
    this._defsCounter = 0;

    // Deterministic per-slide shape data-id generation: "<slideId>-<n>" (e.g. "1-1", "1_1-2").
    // Set (prefix + counter reset) before each convertSlideToHtml call; increments per shape group.
    this._currentSlideDataIdPrefix = '';
    this._currentSlideDataIdCounter = 0;
  }
}

// Compose the prototype from the ported method modules (order-independent).
Object.assign(
  PptxConverter.prototype,
  require('./utilities'),
  require('./colorHelper'),
  require('./themeHelper'),
  require('./placeholderResolver'),
  require('./backgroundHelper'),
  require('./fontHelper'),
  require('./animationStageHelper'),
  require('./slideTransitionHelper'),
  require('./textRenderer'),
  require('./shapeRenderer'),
  require('./run')
);

module.exports = { PptxConverter };
