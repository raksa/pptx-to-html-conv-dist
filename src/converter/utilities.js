'use strict';

// Port of PptxToHtml/pptx/Utilities.cs

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { firstChild, elements, attr, childElements, NS } = require('../ooxml/dom');
const { htmlEncode } = require('../util/format');

const isRelNs = (ns) => ns === NS.R || (ns || '').toLowerCase().includes('relationships');

module.exports = {
  emuToPx(emu) {
    if (emu == null) return 0;
    return (emu / 914400.0) * 96.0;
  },

  encode(text) {
    return htmlEncode(text == null ? '' : text);
  },

  getStringValue(value, defaultValue = '') {
    return value == null ? defaultValue : value;
  },

  resolveImagePart(partContainer, blip) {
    const relId = this.getSvgBlipEmbedId(blip) || (blip ? attr(blip, 'embed', isRelNs) : null);
    if (!relId) return null;
    try {
      return partContainer.getPartById(relId);
    } catch {
      return null;
    }
  },

  getSvgBlipEmbedId(blip) {
    const extList = blip ? firstChild(blip, 'extLst') : null;
    if (!extList) return null;

    for (const ext of elements(extList, 'ext')) {
      for (const child of childElements(ext)) {
        if (child.localName.toLowerCase() !== 'svgblip') continue;
        for (let i = 0; i < child.attributes.length; i++) {
          const a = child.attributes[i];
          if (a.localName.toLowerCase() !== 'embed') continue;
          const ns = a.namespaceURI || '';
          if (ns === NS.R || ns.toLowerCase().includes('relationships')) {
            return a.value;
          }
        }
      }
    }
    return null;
  },

  saveImageFile(part, trackCurrentSlide = true) {
    const partUri = part.uri;
    if (this._savedImagePaths.has(partUri)) {
      const cached = this._savedImagePaths.get(partUri);
      if (trackCurrentSlide) this._currentSlideImages.push(cached);
      return cached;
    }

    const bytes = Buffer.from(part.getBytes());

    const mime = part.contentType || 'image/png';
    let ext;
    switch (mime) {
      case 'image/jpeg': ext = '.jpg'; break;
      case 'image/gif': ext = '.gif'; break;
      case 'image/bmp': ext = '.bmp'; break;
      case 'image/svg+xml': ext = '.svg'; break;
      case 'image/tiff': ext = '.tiff'; break;
      case 'image/webp': ext = '.webp'; break;
      default: ext = '.png';
    }

    this._imageCounter++;
    const fileName = `image${this._imageCounter}${ext}`;
    const filePath = path.join(this._outputDir, 'images', fileName);
    fs.writeFileSync(filePath, bytes);

    const relativePath = `images/${fileName}`;
    this._savedImagePaths.set(partUri, relativePath);
    if (trackCurrentSlide) this._currentSlideImages.push(relativePath);
    return relativePath;
  },

  saveMediaFile(dataPart) {
    const partUri = dataPart.uri;
    const mime = dataPart.contentType || 'application/octet-stream';

    if (this._savedMediaPaths.has(partUri)) {
      const cached = this._savedMediaPaths.get(partUri);
      this.trackMediaByType(mime, cached);
      return cached;
    }

    const bytes = Buffer.from(dataPart.getBytes());

    let ext;
    switch (mime) {
      case 'video/mp4': ext = '.mp4'; break;
      case 'video/webm': ext = '.webm'; break;
      case 'video/ogg': ext = '.ogv'; break;
      case 'video/x-msvideo': ext = '.avi'; break;
      case 'video/x-ms-wmv': ext = '.wmv'; break;
      case 'audio/mpeg': ext = '.mp3'; break;
      case 'audio/wav':
      case 'audio/x-wav': ext = '.wav'; break;
      case 'audio/ogg': ext = '.ogg'; break;
      case 'audio/mp4':
      case 'audio/x-m4a': ext = '.m4a'; break;
      default: {
        const e = path.extname(partUri);
        ext = e && e.length > 0 ? e : '.bin';
      }
    }

    this._mediaCounter++;
    const fileName = `media${this._mediaCounter}${ext}`;
    const filePath = path.join(this._outputDir, 'media', fileName);
    fs.writeFileSync(filePath, bytes);

    const relativePath = `media/${fileName}`;
    this._savedMediaPaths.set(partUri, relativePath);
    this.trackMediaByType(mime, relativePath);
    return relativePath;
  },

  trackMediaByType(mime, filePath) {
    if (mime.toLowerCase().startsWith('video/')) this._currentSlideVideos.push(filePath);
    else if (mime.toLowerCase().startsWith('audio/')) this._currentSlideAudios.push(filePath);
  },

  writeChecksumFile(inputPath, outputDir) {
    const fileBytes = fs.readFileSync(inputPath);
    const md5 = crypto.createHash('md5').update(fileBytes).digest('hex').toUpperCase();
    const sha256 = crypto.createHash('sha256').update(fileBytes).digest('hex').toUpperCase();
    const checksumPath = path.join(outputDir, 'checksum.txt');
    fs.writeFileSync(
      checksumPath,
      [
        `File: ${path.basename(inputPath)}`,
        `MD5: ${md5}`,
        `SHA256: ${sha256}`,
        '',
      ].join('\n'),
      'utf8'
    );
    return checksumPath;
  },
};
