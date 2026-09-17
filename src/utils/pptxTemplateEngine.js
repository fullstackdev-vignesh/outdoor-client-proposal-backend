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

    const newBaseName = `slide_gen_${this._nextSlideIndex++}`;
    const newSlidePath = `ppt/slides/${newBaseName}.xml`;
    const newRelsPath = `ppt/slides/_rels/${newBaseName}.xml.rels`;
    this.writeText(newSlidePath, slideXml);
    this.writeText(newRelsPath, relsXml);

    // Register relationship and content type for the cloned slide
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
