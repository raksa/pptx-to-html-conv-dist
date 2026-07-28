'use strict';

// Port of PptxToHtml/pptx/PlaceholderResolver.cs

const { firstChild, attr } = require('../ooxml/dom');
const { slideLayoutPart, slideMasterPart } = require('../ooxml/parts');
const { placeholderOf, textBodyOf, shapeTreeOf, hasChildElements } = require('./elements');

const LEVEL_NAMES = [
  'lvl1pPr', 'lvl2pPr', 'lvl3pPr', 'lvl4pPr', 'lvl5pPr',
  'lvl6pPr', 'lvl7pPr', 'lvl8pPr', 'lvl9pPr',
];

const BODY_STYLE_TYPES = new Set(['body', 'subTitle', 'obj', 'clipArt', 'tbl', 'chart', 'media']);
const TITLE_STYLE_TYPES = new Set(['title', 'ctrTitle']);

// Holds inherited text properties resolved from the placeholder chain.
function makeInherited(runProps, alignment, verticalAnchor, bulletParaProps, bulletStyleSource, fallbackStyleSource) {
  return { runProps, alignment, verticalAnchor, bulletParaProps, bulletStyleSource, fallbackStyleSource };
}

function anchorToken(av) {
  if (av === 'b') return 'b';
  if (av === 'ctr') return 'ctr';
  return 't';
}

module.exports = {
  makeInherited,

  masterTextStyleFor(txStyles, phType) {
    if (TITLE_STYLE_TYPES.has(phType)) return firstChild(txStyles, 'titleStyle');
    if (BODY_STYLE_TYPES.has(phType)) return firstChild(txStyles, 'bodyStyle');
    return firstChild(txStyles, 'otherStyle');
  },

  resolveInheritedProps(ph, slidePart) {
    if (ph == null) return makeInherited(null, null, null, null, null, null);

    const layoutPart = slideLayoutPart(slidePart);
    if (layoutPart == null) return makeInherited(null, null, null, null, null, null);

    let resolvedRunProps = null;
    let resolvedAlign = null;
    let resolvedAnchor = null;
    let resolvedBulletProps = null;

    const layoutMatch = this.findMatchingPlaceholder(ph, shapeTreeOf(layoutPart.root));

    let effectivePhType = attr(ph, 'type');
    if (effectivePhType == null && layoutMatch != null) {
      effectivePhType = attr(placeholderOf(layoutMatch), 'type');
    }

    if (layoutMatch != null) {
      const bodyPr = firstChild(textBodyOf(layoutMatch), 'bodyPr');
      const av = attr(bodyPr, 'anchor');
      if (av != null) resolvedAnchor = anchorToken(av);

      const layoutDefRp = this.getDefRunPropsFromShape(layoutMatch, 0);
      if (layoutDefRp != null && this.hasUsefulProps(layoutDefRp)) resolvedRunProps = layoutDefRp;

      const layoutL1 = firstChild(firstChild(textBodyOf(layoutMatch), 'lstStyle'), 'lvl1pPr');
      const l1Align = attr(layoutL1, 'algn');
      if (l1Align != null) resolvedAlign = l1Align;

      if (firstChild(layoutL1, 'buNone') != null) resolvedBulletProps = null;
      else if (firstChild(layoutL1, 'buChar') != null) resolvedBulletProps = layoutL1;
    }

    const layoutSuppressesBullets =
      resolvedBulletProps == null &&
      layoutMatch != null &&
      firstChild(firstChild(firstChild(textBodyOf(layoutMatch), 'lstStyle'), 'lvl1pPr'), 'buNone') != null;

    const masterPart = slideMasterPart(layoutPart);
    if (masterPart == null) {
      const earlyLayoutLs = layoutMatch ? firstChild(textBodyOf(layoutMatch), 'lstStyle') : null;
      const earlyBulletStyleSource =
        earlyLayoutLs != null && hasChildElements(earlyLayoutLs) ? earlyLayoutLs : null;
      return makeInherited(resolvedRunProps, resolvedAlign, resolvedAnchor, null, earlyBulletStyleSource, null);
    }

    const masterRoot = masterPart.root;
    const masterMatch = this.findMatchingPlaceholder(ph, shapeTreeOf(masterRoot));
    if (resolvedRunProps == null) {
      const masterDefRp = this.getDefRunPropsFromShape(masterMatch, 0);
      if (masterDefRp != null && this.hasUsefulProps(masterDefRp)) resolvedRunProps = masterDefRp;
    }

    if (resolvedAnchor == null && masterMatch != null) {
      const masterBodyPr = firstChild(textBodyOf(masterMatch), 'bodyPr');
      const av = attr(masterBodyPr, 'anchor');
      if (av != null) resolvedAnchor = anchorToken(av);
    }

    if (effectivePhType == null && masterMatch != null) {
      effectivePhType = attr(placeholderOf(masterMatch), 'type');
    }

    const txStyles = firstChild(masterRoot, 'txStyles');
    if (txStyles != null) {
      const style = this.masterTextStyleFor(txStyles, effectivePhType);
      if (style != null) {
        const l1 = firstChild(style, 'lvl1pPr');
        if (l1 != null) {
          if (resolvedRunProps == null) resolvedRunProps = firstChild(l1, 'defRPr');
          if (resolvedAlign == null && attr(l1, 'algn') != null) resolvedAlign = attr(l1, 'algn');
          if (!layoutSuppressesBullets && resolvedBulletProps == null && firstChild(l1, 'buChar') != null) {
            resolvedBulletProps = l1;
          }
        }
      }
    }

    let masterTextStyle = null;
    if (txStyles != null) masterTextStyle = this.masterTextStyleFor(txStyles, effectivePhType);

    let bulletStyleSource = null;
    let fallbackStyleSource = null;
    const layoutLs = layoutMatch ? firstChild(textBodyOf(layoutMatch), 'lstStyle') : null;
    if (layoutLs != null && hasChildElements(layoutLs)) {
      bulletStyleSource = layoutLs;
      fallbackStyleSource = masterTextStyle;
    } else {
      bulletStyleSource = masterTextStyle;
    }

    return makeInherited(
      resolvedRunProps,
      resolvedAlign,
      resolvedAnchor,
      resolvedBulletProps,
      bulletStyleSource,
      fallbackStyleSource
    );
  },

  getDefRunPropsFromShape(shape, level) {
    const ls = firstChild(textBodyOf(shape), 'lstStyle');
    if (ls == null) return null;
    const levelProps = this.getLevelParagraphProperties(ls, level);
    return firstChild(levelProps, 'defRPr');
  },

  getLevelParagraphProperties(lstStyle, level) {
    return this.getLevelPropsFromComposite(lstStyle, level);
  },

  getLevelPropsFromComposite(element, level) {
    if (element == null) return null;
    const name = LEVEL_NAMES[level] || LEVEL_NAMES[0];
    return firstChild(element, name);
  },

  hasUsefulProps(drp) {
    return (
      attr(drp, 'sz') != null ||
      firstChild(drp, 'solidFill') != null ||
      attr(drp, 'b') != null ||
      attr(drp, 'i') != null ||
      firstChild(drp, 'latin') != null ||
      firstChild(drp, 'cs') != null ||
      firstChild(drp, 'ea') != null
    );
  },

  getUsefulDefRunProps(levelProps) {
    const drp = firstChild(levelProps, 'defRPr');
    if (drp != null && this.hasUsefulProps(drp)) return drp;
    return null;
  },
};
