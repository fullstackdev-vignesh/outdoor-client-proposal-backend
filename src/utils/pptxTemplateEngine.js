const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const MASTER_PPTX_PATH = path.join(__dirname, '..', '..', 'assets', 'proposal-templates', 'master.pptx');

// Slide 1 = cover, Slide 2/3 = static About Us / Why Choose Us, Slide 4 = per-site
// "Media Specifications" template, Slide 5 = per-site map template, Slide 6 = Thank You.
const SITE_SPEC_TEMPLATE = 'slide4';
const SITE_MAP_TEMPLATE = 'slide5';
const THANK_YOU_SLIDE = 'slide6';

function xmlEscape(str) {
  return String(str ?? '').replace(/[<>&'"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' }[c]));
}

function replaceRunOnce(xml, oldText, newText) {
  const target = `<a:t>${oldText}</a:t>`;
  if (!xml.includes(target)) return xml;
  return xml.replace(target, `<a:t>${xmlEscape(newText)}</a:t>`);
}

function extOf(filename) {
  return path.extname(filename).replace('.', '').toLowerCase();
}

class PptxTemplate {
  constructor(zip) {
    this.zip = zip;
    this._nextMediaIndex = 9000;
    this._nextSlideIndex = 9000;
    this._nextRelId = 9000;
    this._nextSldId = 900;
  }

  static async load() {
    const buf = fs.readFileSync(MASTER_PPTX_PATH);
    const zip = await JSZip.loadAsync(buf);
    return new PptxTemplate(zip);
  }

  async readText(partPath) {
    return this.zip.file(partPath).async('string');
  }

  writeText(partPath, content) {
    this.zip.file(partPath, content);
  }

  async addMediaFile(buffer, ext) {
    const name = `image_gen_${this._nextMediaIndex++}.${ext}`;
    this.zip.file(`ppt/media/${name}`, buffer);
    return `../media/${name}`;
  }

  /** Replace the media Target for a given relationship Id inside a slide's .rels XML. */
  replaceRelTarget(relsXml, relId, newTarget) {
    const re = new RegExp(`(<Relationship Id="${relId}"[^>]*Target=")[^"]*(")`);
    return relsXml.replace(re, `$1${newTarget}$2`);
  }

  async cloneSlide(templateBaseName, mutations) {
    const slidePath = `ppt/slides/${templateBaseName}.xml`;
    const relsPath = `ppt/slides/_rels/${templateBaseName}.xml.rels`;

    let slideXml = await this.readText(slidePath);
    let relsXml = await this.readText(relsPath);

    for (const [oldText, newText] of mutations.textReplacements || []) {
      slideXml = replaceRunOnce(slideXml, oldText, newText);
    }

    for (const { relId, buffer, ext } of mutations.imageReplacements || []) {
      if (!buffer) continue;
      const target = await this.addMediaFile(buffer, ext);
      relsXml = this.replaceRelTarget(relsXml, relId, target);
    }

    const newBaseName = `slide_gen_${this._nextSlideIndex++}`;
    const newSlidePath = `ppt/slides/${newBaseName}.xml`;
    const newRelsPath = `ppt/slides/_rels/${newBaseName}.xml.rels`;
    this.writeText(newSlidePath, slideXml);
    this.writeText(newRelsPath, relsXml);

    return newBaseName;
  }

  async setCoverFields({ customerLabel, dateLabel }) {
    let slideXml = await this.readText('ppt/slides/slide1.xml');
    slideXml = replaceRunOnce(slideXml, 'Maxi Vision Eye Hospital', customerLabel);
    slideXml = replaceRunOnce(slideXml, 'ate: Aug 10, 2026', `ate: ${dateLabel}`);
    this.writeText('ppt/slides/slide1.xml', slideXml);
  }

  /** Registers newly cloned slide parts into presentation.xml / rels / [Content_Types].xml, inserted after `afterBaseName`. */
  async insertSlides(baseNames, afterBaseName) {
    const contentTypesPath = '[Content_Types].xml';
    let contentTypes = await this.readText(contentTypesPath);
    let presRels = await this.readText('ppt/_rels/presentation.xml.rels');
    let presentation = await this.readText('ppt/presentation.xml');

    const newSldIdEntries = [];
    for (const baseName of baseNames) {
      contentTypes = contentTypes.replace(
        '</Types>',
        `<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml" PartName="/ppt/slides/${baseName}.xml"/></Types>`
      );

      const relId = `rIdGen${this._nextRelId++}`;
      presRels = presRels.replace(
        '</Relationships>',
        `<Relationship Id="${relId}" Target="slides/${baseName}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/></Relationships>`
      );

      const sldId = this._nextSldId++;
      newSldIdEntries.push(`<p:sldId id="${sldId}" r:id="${relId}"/>`);
    }

    // Find the r:id of the "afterBaseName" slide part so we can splice the new sldId entries right after it.
    const afterRelMatch = presRels.match(new RegExp(`Id="(rId\\d+)" Target="slides/${afterBaseName}\\.xml"`));
    const afterRelId = afterRelMatch ? afterRelMatch[1] : null;

    if (afterRelId) {
      const anchorRe = new RegExp(`(<p:sldId[^>]*r:id="${afterRelId}"/>)`);
      presentation = presentation.replace(anchorRe, `$1${newSldIdEntries.join('')}`);
    } else {
      presentation = presentation.replace('</p:sldIdLst>', `${newSldIdEntries.join('')}</p:sldIdLst>`);
    }

    this.writeText(contentTypesPath, contentTypes);
    this.writeText('ppt/_rels/presentation.xml.rels', presRels);
    this.writeText('ppt/presentation.xml', presentation);
  }

  /** Removes a slide's <p:sldId> entry from the slide order (part stays in the zip, just unreferenced/unused). */
  async removeFromSlideOrder(baseName) {
    let presRels = await this.readText('ppt/_rels/presentation.xml.rels');
    let presentation = await this.readText('ppt/presentation.xml');

    const relMatch = presRels.match(new RegExp(`Id="(rId\\d+)" Target="slides/${baseName}\\.xml"`));
    if (!relMatch) return;
    const relId = relMatch[1];
    presentation = presentation.replace(new RegExp(`<p:sldId[^>]*r:id="${relId}"/>`), '');
    this.writeText('ppt/presentation.xml', presentation);
  }

  async save() {
    return this.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  }
}

module.exports = { PptxTemplate, SITE_SPEC_TEMPLATE, SITE_MAP_TEMPLATE, THANK_YOU_SLIDE, extOf };
