const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const MASTER_PPTX_PATH = path.join(__dirname, '..', '..', 'assets', 'proposal-templates', 'master.pptx');
const BACKEND_ROOT = path.join(__dirname, '..', '..');

function xmlEscape(str) {
  return String(str ?? '').replace(/[<>&'"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' }[c]));
}

function extOf(filename) {
  return path.extname(filename).replace('.', '').toLowerCase();
}

// Minimal, dependency-free width/height reader for the two formats site photos use in practice.
function getImageDimensions(buffer, ext) {
  try {
    if (/^png$/i.test(ext) && buffer.length >= 24 && buffer.readUInt32BE(0) === 0x89504e47) {
      return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
    }
    if (/^jpe?g$/i.test(ext) && buffer[0] === 0xff && buffer[1] === 0xd8) {
      let i = 2;
      while (i + 9 < buffer.length) {
        if (buffer[i] !== 0xff) {
          i++;
          continue;
        }
        const marker = buffer[i + 1];
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
          return { height: buffer.readUInt16BE(i + 5), width: buffer.readUInt16BE(i + 7) };
        }
        i += 2 + buffer.readUInt16BE(i + 2);
      }
    }
  } catch {
    // fall through to null below
  }
  return null;
}

// "object-fit: cover" style crop, expressed as OOXML <a:srcRect> insets (thousandths of a percent).
function computeCoverFillRect(srcW, srcH, boxW, boxH) {
  if (!srcW || !srcH || !boxW || !boxH) return { l: 0, t: 0, r: 0, b: 0 };
  const srcAR = srcW / srcH;
  const boxAR = boxW / boxH;
  if (srcAR > boxAR) {
    const crop = Math.round(((1 - boxAR / srcAR) / 2) * 100000);
    return { l: crop, t: 0, r: crop, b: 0 };
  }
  if (srcAR < boxAR) {
    const crop = Math.round(((1 - srcAR / boxAR) / 2) * 100000);
    return { l: 0, t: crop, r: 0, b: crop };
  }
  return { l: 0, t: 0, r: 0, b: 0 };
}

function removeRedHighlightShapes(slideXml) {
  const spRegex = /<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g;
  return slideXml.replace(spRegex, (match) => {
    if (/srgbClr\s+val="(?:FF0000|C00000|ED1C24|FF0022|990000)"/i.test(match)) {
      return '';
    }
    return match;
  });
}

class PptxTemplate {
  constructor(zip) {
    this.zip = zip;
    this._nextMediaIndex = 9000;
    this._nextSlideIndex = 9000;
    this._nextRelId = 9000;
    this._nextSldId = 900;
  }

  static async load(customPath) {
    let buf;
    if (customPath && /^https?:\/\//i.test(customPath)) {
      try {
        const response = await fetch(customPath);
        if (!response.ok) {
          throw new Error(`Failed to fetch template from URL: ${customPath} (${response.status})`);
        }
        const arrayBuffer = await response.arrayBuffer();
        buf = Buffer.from(arrayBuffer);
      } catch (err) {
        console.warn('Failed to fetch remote template, falling back to master.pptx:', err.message);
        buf = fs.readFileSync(MASTER_PPTX_PATH);
      }
    } else {
      let targetPath = MASTER_PPTX_PATH;
      if (customPath) {
        const absCustom = path.resolve(BACKEND_ROOT, customPath.replace(/^\//, ''));
        if (fs.existsSync(absCustom)) {
          targetPath = absCustom;
        }
      }
      buf = fs.existsSync(targetPath) ? fs.readFileSync(targetPath) : fs.readFileSync(MASTER_PPTX_PATH);
    }

    const zip = await JSZip.loadAsync(buf);
    const tpl = new PptxTemplate(zip);
    await tpl.ensureImageDefaults();
    return tpl;
  }

  async readText(partPath) {
    return this.zip.file(partPath).async('string');
  }

  writeText(partPath, content) {
    this.zip.file(partPath, content);
  }

  async ensureImageDefaults() {
    const contentTypesPath = '[Content_Types].xml';
    if (!this.zip.file(contentTypesPath)) return;
    let contentTypes = await this.readText(contentTypesPath);
    const defaults = [
      '<Default Extension="png" ContentType="image/png"/>',
      '<Default Extension="jpeg" ContentType="image/jpeg"/>',
      '<Default Extension="jpg" ContentType="image/jpeg"/>',
      '<Default Extension="webp" ContentType="image/webp"/>',
    ];
    for (const d of defaults) {
      if (!contentTypes.includes(d)) {
        contentTypes = contentTypes.replace('</Types>', `${d}</Types>`);
      }
    }
    this.writeText(contentTypesPath, contentTypes);
  }

  async getSlideFiles() {
    const files = [];
    const slidesFolder = this.zip.folder('ppt/slides');
    if (slidesFolder) {
      slidesFolder.forEach((relativePath) => {
        if (/^slide\d+\.xml$/.test(relativePath)) {
          files.push(relativePath.replace('.xml', ''));
        }
      });
    }
    files.sort((a, b) => parseInt(a.replace('slide', '')) - parseInt(b.replace('slide', '')));
    return files;
  }

  async addMediaFile(buffer, ext) {
    const name = `image_gen_${this._nextMediaIndex++}.${ext}`;
    this.zip.file(`ppt/media/${name}`, buffer);
    return `../media/${name}`;
  }

  replaceRelTarget(relsXml, relId, newTarget) {
    const re = new RegExp(`(<Relationship Id="${relId}"[^>]*Target=")[^"]*(")`);
    return relsXml.replace(re, `$1${newTarget}$2`);
  }

  async cloneSlide(templateBaseName, mutations, siteContext) {
    const slidePath = `ppt/slides/${templateBaseName}.xml`;
    const relsPath = `ppt/slides/_rels/${templateBaseName}.xml.rels`;

    let slideXml = await this.readText(slidePath);
    let relsXml = this.zip.file(relsPath) ? await this.readText(relsPath) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

    // Remove red highlight boxes from template slide
    slideXml = removeRedHighlightShapes(slideXml);

    // 1. Exact text replacements
    for (const [oldText, newText] of mutations.textReplacements || []) {
      const target = `<a:t>${oldText}</a:t>`;
      if (slideXml.includes(target)) {
        slideXml = slideXml.replace(target, `<a:t>${xmlEscape(newText)}</a:t>`);
      }
    }

    // 2. Smart fallback replacement for location/description if siteContext is provided
    if (siteContext) {
      const siteDesc = `${siteContext.location || siteContext.mediaName || siteContext.city} (${siteContext.mediaId || '-'}) ${siteContext.width && siteContext.height ? `${siteContext.width}x${siteContext.height}` : ''}`.trim();
      const locationTagRegex = /<a:t>([^<]*(?:flyover|bridge|towards|junction|bypass|signal|road|street|avenue|cross|sq|st|pass)[^<]*)<\/a:t>/gi;
      if (locationTagRegex.test(slideXml)) {
        slideXml = slideXml.replace(locationTagRegex, `<a:t>${xmlEscape(siteDesc)}</a:t>`);
      }
    }

    // 3. Image replacements
    for (const { buffer, ext } of mutations.imageReplacements || []) {
      if (!buffer) continue;
      const target = await this.addMediaFile(buffer, ext);
      const imageRelMatches = [...relsXml.matchAll(/Id="(rId\d+)"[^>]*Type="[^"]*relationships\/image"/g)];
      if (imageRelMatches.length > 0) {
        for (const match of imageRelMatches) {
          const relId = match[1];
          relsXml = this.replaceRelTarget(relsXml, relId, target);
        }
      } else {
        relsXml = this.replaceRelTarget(relsXml, 'rId2', target);
      }
    }

    return this._registerClonedSlide(slideXml, relsXml);
  }

  // Shared by cloneSlide/cloneAdinnSiteSlide: writes the new slide+rels parts and registers
  // them in [Content_Types].xml and presentation.xml.rels so they show up as real slides.
  async _registerClonedSlide(slideXml, relsXml) {
    const newBaseName = `slide_gen_${this._nextSlideIndex++}`;
    const newSlidePath = `ppt/slides/${newBaseName}.xml`;
    const newRelsPath = `ppt/slides/_rels/${newBaseName}.xml.rels`;
    this.writeText(newSlidePath, slideXml);
    this.writeText(newRelsPath, relsXml);

    const newRelId = `rIdGen${this._nextRelId++}`;
    const contentTypesPath = '[Content_Types].xml';
    let contentTypes = await this.readText(contentTypesPath);
    const overrideTag = `PartName="/ppt/slides/${newBaseName}.xml"`;
    if (!contentTypes.includes(overrideTag)) {
      contentTypes = contentTypes.replace(
        '</Types>',
        `<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml" PartName="/ppt/slides/${newBaseName}.xml"/></Types>`
      );
      this.writeText(contentTypesPath, contentTypes);
    }

    let presRels = await this.readText('ppt/_rels/presentation.xml.rels');
    presRels = presRels.replace(
      '</Relationships>',
      `<Relationship Id="${newRelId}" Target="slides/${newBaseName}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/></Relationships>`
    );
    this.writeText('ppt/_rels/presentation.xml.rels', presRels);

    return newBaseName;
  }

  // adinn-new-template slide 4/5 cloning: exact text swaps, plus per-relationship-id image
  // replacement that also recomputes that shape's <a:fillRect> crop so the new photo covers
  // its existing box without stretching (aspect-fit "cover", not distort-to-fill).
  // images: [{ relId, buffer, ext, boxWidthEMU, boxHeightEMU }]
  // clearImageRelId: relationship id whose blipFill should become empty (map placeholder)
  // placeholderText: { offX, offY, extCx, extCy, text } — new centered textbox for that empty area
  async cloneAdinnSiteSlide(templateBaseName, { textReplacements = [], images = [], clearImageRelId, placeholderText, boxWidths = [] } = {}) {
    const slidePath = `ppt/slides/${templateBaseName}.xml`;
    const relsPath = `ppt/slides/_rels/${templateBaseName}.xml.rels`;

    let slideXml = await this.readText(slidePath);
    let relsXml = await this.readText(relsPath);

    for (const [oldText, newText] of textReplacements) {
      const target = `<a:t>${oldText}</a:t>`;
      if (slideXml.includes(target)) {
        slideXml = slideXml.replace(target, `<a:t>${xmlEscape(newText)}</a:t>`);
      }
    }

    // Widen a specific textbox (matched by its exact <a:off>) so its text fits on one line
    // instead of wrapping character-by-character. Only cx changes — position/height/style stay.
    for (const { offX, offY, widthEMU } of boxWidths) {
      const re = new RegExp(`(<a:off x="${offX}" y="${offY}"\\/><a:ext cx=")\\d+("\\s*cy="\\d+"\\/>)`);
      slideXml = slideXml.replace(re, `$1${widthEMU}$2`);
    }

    for (const { relId, buffer, ext, boxWidthEMU, boxHeightEMU } of images) {
      if (!buffer) continue;
      const newTarget = await this.addMediaFile(buffer, ext);
      relsXml = this.replaceRelTarget(relsXml, relId, newTarget);

      // <a:fillRect> insets shrink the visible image *within* the shape (leaving gaps) — the
      // correct "cover crop" element is <a:srcRect>, which crops the source image itself before
      // the (gap-free) stretch-to-fill-shape happens.
      const dims = getImageDimensions(buffer, ext);
      const rect = dims ? computeCoverFillRect(dims.width, dims.height, boxWidthEMU, boxHeightEMU) : { l: 0, t: 0, r: 0, b: 0 };
      const blipStretchRe = new RegExp(`<a:blip r:embed="${relId}"\\/><a:stretch><a:fillRect[^/]*\\/><\\/a:stretch>`);
      slideXml = slideXml.replace(
        blipStretchRe,
        `<a:blip r:embed="${relId}"/><a:srcRect l="${rect.l}" t="${rect.t}" r="${rect.r}" b="${rect.b}"/><a:stretch><a:fillRect/></a:stretch>`
      );
    }

    if (clearImageRelId) {
      const clearRe = new RegExp(`<a:blipFill><a:blip r:embed="${clearImageRelId}"\\/><a:stretch>.*?<\\/a:stretch><\\/a:blipFill>`);
      slideXml = slideXml.replace(clearRe, '<a:noFill/>');
    }

    if (placeholderText) {
      const { offX, offY, extCx, extCy, text } = placeholderText;
      const shapeId = 9000 + this._nextSlideIndex;
      const placeholderSp =
        `<p:sp><p:nvSpPr><p:cNvPr name="Map Placeholder" id="${shapeId}"/><p:cNvSpPr txBox="true"/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${extCx}" cy="${extCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
        `<p:txBody><a:bodyPr anchor="ctr" wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
        `<a:r><a:rPr lang="en-US" sz="2400"><a:solidFill><a:srgbClr val="666666"/></a:solidFill><a:latin typeface="Times New Roman MT"/></a:rPr><a:t>${xmlEscape(text)}</a:t></a:r>` +
        `</a:p></p:txBody></p:sp>`;
      slideXml = slideXml.replace('</p:spTree>', `${placeholderSp}</p:spTree>`);
    }

    return this._registerClonedSlide(slideXml, relsXml);
  }

  async setCoverFields({ customerLabel, dateLabel }) {
    let slideXml = await this.readText('ppt/slides/slide1.xml');
    slideXml = slideXml.replace(/<a:t>[^<]*<\/a:t>/g, (match) => {
      if (match.includes('Maxi Vision') || match.includes('Hospital') || match.includes('Proposal')) {
        return `<a:t>${xmlEscape(customerLabel)}</a:t>`;
      }
      return match;
    });
    this.writeText('ppt/slides/slide1.xml', slideXml);
  }

  // adinn-new-template's slide1 splits "Date: Aug 10, 2026" across two runs ("D" + "ate: Aug 10, 2026").
  // Only the "ate: ..." run is touched so the "D" run/formatting is left completely intact.
  async setCoverDateLabel(dateLabel) {
    let slideXml = await this.readText('ppt/slides/slide1.xml');
    slideXml = slideXml.replace(/<a:t>ate:\s*[^<]*<\/a:t>/, `<a:t>ate: ${xmlEscape(dateLabel)}</a:t>`);
    this.writeText('ppt/slides/slide1.xml', slideXml);
  }

  // Widens (and optionally re-centers) a single textbox on a given slide part, matched by its
  // exact original <a:off>. Only x/width change — y, height, font, color and every other
  // element on the slide are left untouched.
  async resizeTextBox(slidePath, { offX, offY, newOffX = offX, newWidthEMU }) {
    let slideXml = await this.readText(slidePath);
    const re = new RegExp(`<a:off x="${offX}" y="${offY}"\\/><a:ext cx="\\d+"( cy="\\d+"\\/>)`);
    slideXml = slideXml.replace(re, `<a:off x="${newOffX}" y="${offY}"/><a:ext cx="${newWidthEMU}"$1`);
    this.writeText(slidePath, slideXml);
  }

  async setFinalSlideOrder(orderedBaseNames) {
    const presRelsXml = await this.readText('ppt/_rels/presentation.xml.rels');
    let presentationXml = await this.readText('ppt/presentation.xml');

    const relEntries = [...presRelsXml.matchAll(/<Relationship[^>]*>/g)];
    const targetToRelId = {};
    for (const match of relEntries) {
      const tag = match[0];
      const idMatch = tag.match(/Id="([^"]+)"/);
      const targetMatch = tag.match(/Target="([^"]+)"/);
      if (idMatch && targetMatch) {
        const id = idMatch[1];
        const target = targetMatch[1];
        const baseNameMatch = target.match(/slides\/([^/.]+)\.xml$/);
        if (baseNameMatch) {
          targetToRelId[baseNameMatch[1]] = id;
        }
      }
    }

    const sldIdEntries = [];
    for (const baseName of orderedBaseNames) {
      const relId = targetToRelId[baseName];
      if (relId) {
        const sldId = this._nextSldId++;
        sldIdEntries.push(`<p:sldId id="${sldId}" r:id="${relId}"/>`);
      }
    }

    presentationXml = presentationXml.replace(
      /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
      `<p:sldIdLst>${sldIdEntries.join('')}</p:sldIdLst>`
    );

    this.writeText('ppt/presentation.xml', presentationXml);
  }

  async removeFromSlideOrder(baseName) {
    let presRels = await this.readText('ppt/_rels/presentation.xml.rels');
    let presentation = await this.readText('ppt/presentation.xml');

    const relMatch = presRels.match(new RegExp(`Id="(rId\\d+)"[^>]*Target="(?:/ppt/)?slides/${baseName}\\.xml"`)) ||
                     presRels.match(new RegExp(`Target="(?:/ppt/)?slides/${baseName}\\.xml"[^>]*Id="(rId\\d+)"`));
    if (!relMatch) return;
    const relId = relMatch[1];
    presentation = presentation.replace(new RegExp(`<p:sldId[^>]*r:id="${relId}"[^/>]*/>`), '');
    this.writeText('ppt/presentation.xml', presentation);
  }

  async save() {
    return this.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  }
}

module.exports = { PptxTemplate, extOf };
