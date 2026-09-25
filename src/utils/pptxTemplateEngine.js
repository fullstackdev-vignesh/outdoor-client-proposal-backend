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
// Used only for shapes that keep their box's exact size/position (no "contain" repositioning) —
// crops the source image so it fills the box with no letterboxing, at the cost of trimming
// whichever axis overhangs.
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

// Aspect-ratio preserving contain fit (no crop, no stretch). Fits full image in (maxW, maxH) area,
// centered horizontally and vertically.
function computeContainBox(srcW, srcH, areaX, areaY, maxW, maxH) {
  if (!srcW || !srcH || !maxW || !maxH) {
    return { offX: areaX, offY: areaY, extCx: maxW, extCy: maxH };
  }
  const srcAR = srcW / srcH;
  const areaAR = maxW / maxH;
  let fittedW, fittedH, fittedX, fittedY;

  if (srcAR > areaAR) {
    fittedW = maxW;
    fittedH = Math.round(maxW / srcAR);
    fittedX = areaX;
    fittedY = areaY + Math.round((maxH - fittedH) / 2);
  } else {
    fittedH = maxH;
    fittedW = Math.round(maxH * srcAR);
    fittedX = areaX + Math.round((maxW - fittedW) / 2);
    fittedY = areaY;
  }

  return { offX: fittedX, offY: fittedY, extCx: fittedW, extCy: fittedH };
}

// Gap kept between the route-info label ("24 mins • 6.8 km") and the map's own bottom edge, so
// the label reads as floating just above the map's border instead of touching/overlapping it.
const MAP_LABEL_MARGIN_EMU = 130000;

// Box for the route-info label — sized/positioned off the map's actually-visible (fitted) rect,
// with a small bottom margin so it never touches that rect's own bottom edge/border.
function computeMapLabelBox(fittedRect) {
  const labelHeight = Math.min(380000, fittedRect.extCy);
  const margin = Math.min(MAP_LABEL_MARGIN_EMU, Math.max(fittedRect.extCy - labelHeight, 0));
  return {
    offX: fittedRect.offX,
    offY: fittedRect.offY + fittedRect.extCy - labelHeight - margin,
    extCx: fittedRect.extCx,
    extCy: labelHeight,
  };
}

// True when a shape's own visual outline is an (effectively) plain axis-aligned rectangle — a
// real preset rectangle, or a "Freeform" custGeom whose path is nothing but its own bounding box
// drawn as four straight edges (a common artifact of templates authored/exported from design
// tools, where every plain rectangle becomes a custom path). Only shapes like this are safe to
// resize/reposition for a "contain" fit: anything with an actual decorative cut/notch would
// visibly distort if stretched to a new aspect ratio, since custGeom paths scale independently
// per axis to fill whatever <a:ext> they're given.
function isPlainRectangleShape(shapeBlock) {
  if (/<a:prstGeom prst="rect"/.test(shapeBlock)) return true;
  const pathMatch = shapeBlock.match(/<a:path w="(\d+)" h="(\d+)">([\s\S]*?)<\/a:path>/);
  if (!pathMatch) return false;
  const [, w, h, pathBody] = pathMatch;
  const rectPathRe = new RegExp(
    '^<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
      `<a:lnTo><a:pt x="${w}" y="0"/></a:lnTo>` +
      `<a:lnTo><a:pt x="${w}" y="${h}"/></a:lnTo>` +
      `<a:lnTo><a:pt x="0" y="${h}"/></a:lnTo>` +
      '<a:close/>$'
  );
  return rectPathRe.test(pathBody);
}

// A shape nested inside a <p:grpSp> has its own <a:xfrm> expressed in that group's local
// child-coordinate space (chOff/chExt), which the group's outer xfrm then scales/translates onto
// the actual slide — so its raw off/ext numbers are *not* real slide-absolute EMUs whenever the
// group applies any scaling. Resizing it as if they were would misplace/mis-size the image (and,
// if the group scales X/Y unevenly, even "aspect ratio" computed from the raw numbers would be
// wrong). Detected by counting <p:grpSp> open/close tags before the given index.
function isInsideGroup(slideXml, index) {
  const before = slideXml.slice(0, index);
  const opens = (before.match(/<p:grpSp>/g) || []).length;
  const closes = (before.match(/<\/p:grpSp>/g) || []).length;
  return opens > closes;
}

// Locates the <p:pic> or plain-rectangle <p:sp> whose <a:blip> references relId and returns its
// current <a:xfrm> box (offset + extent) plus enough info to splice a replacement in later.
// Reading the box off the slide part itself (rather than trusting a value passed down from the
// caller) means this always reflects the *final* position/size, even after an earlier mutation
// (narrowing to make room for a map, widening to fill freed space, etc.) has already run. Returns
// null (leaving the caller's existing cover/stretch behavior untouched) for anything not safely
// fittable in place — decorative non-rectangular shapes, or shapes nested inside a group.
function findPicXfrm(slideXml, relId, { allowGrouped = false } = {}) {
  const blockRe = /<p:(pic|sp)>[\s\S]*?<\/p:\1>/g;
  let match;
  while ((match = blockRe.exec(slideXml))) {
    if (!match[0].includes(`r:embed="${relId}"`)) continue;
    if (!isPlainRectangleShape(match[0])) continue;
    if (!allowGrouped && isInsideGroup(slideXml, match.index)) continue;
    const xfrmRe = /<a:xfrm[^>]*><a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/;
    const xfrmMatch = match[0].match(xfrmRe);
    if (!xfrmMatch) continue;
    return {
      blockIndex: match.index,
      block: match[0],
      xfrmTag: xfrmMatch[0],
      offX: parseInt(xfrmMatch[1], 10),
      offY: parseInt(xfrmMatch[2], 10),
      cx: parseInt(xfrmMatch[3], 10),
      cy: parseInt(xfrmMatch[4], 10),
    };
  }
  return null;
}

// Resizes/repositions an existing <p:pic>'s own <a:xfrm> so `dims` (the image just placed inside
// it) displays fully "contain"-fit within that pic's current box — no crop, no stretch, centered.
// This is what lets two images placed side by side (a site photo + its location map) both show
// completely instead of one/both being distort-stretched to exactly fill their box.
function fitPicIntoBox(slideXml, relId, dims) {
  if (!dims) return slideXml;
  const found = findPicXfrm(slideXml, relId);
  if (!found) return slideXml;
  const fitted = computeContainBox(dims.width, dims.height, found.offX, found.offY, found.cx, found.cy);
  const newXfrmTag = `<a:xfrm><a:off x="${fitted.offX}" y="${fitted.offY}"/><a:ext cx="${fitted.extCx}" cy="${fitted.extCy}"/>`;
  const updatedBlock = found.block.replace(found.xfrmTag, newXfrmTag);
  return slideXml.slice(0, found.blockIndex) + updatedBlock + slideXml.slice(found.blockIndex + found.block.length);
}

function removeRedHighlightShapes(slideXml) {
  const spRegex = /<p:(sp|cxnSp)\b[^>]*>[\s\S]*?<\/p:\1>/g;
  return slideXml.replace(spRegex, (match) => {
    if (
      /srgbClr\s+val="(?:FF0000|C00000|ED1C24|FF0022|990000)"/i.test(match) ||
      /prst="downArrow"/i.test(match) ||
      /name="Arrow:\s*Down/i.test(match)
    ) {
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

    // Remove red highlight boxes from template slide. Some templates (e.g. Jagran-template-two's
    // city-divider slide) use red as an actual text color rather than an annotation overlay, so
    // that shape must be preserved — mutations.preserveRedShapes opts out for those slides only.
    if (!mutations.preserveRedShapes) {
      slideXml = removeRedHighlightShapes(slideXml);
    }

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

  updateLocationHyperlink(slideXml, relsXml, locationUrl) {
    if (!locationUrl) return { slideXml, relsXml };

    const escapedUrl = xmlEscape(locationUrl);

    // 1. Replace Target in any existing hyperlink Relationship element in relsXml
    const hasHyperlinkRel = /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"/i.test(relsXml);

    if (hasHyperlinkRel) {
      relsXml = relsXml
        .replace(
          /(<Relationship\b[^>]*?)\bTarget="[^"]*"([^>]*?Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"[^>]*\/>)/gi,
          `$1Target="${escapedUrl}"$2`
        )
        .replace(
          /(<Relationship\b[^>]*?Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/hyperlink"[^>]*?)\bTarget="[^"]*"([^>]*\/>)/gi,
          `$1Target="${escapedUrl}"$2`
        );
    } else {
      const relId = `rIdHlinkLoc${this._nextRelId++}`;
      const newRel = `<Relationship Id="${relId}" Target="${escapedUrl}" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"/>`;
      relsXml = relsXml.replace('</Relationships>', `${newRel}</Relationships>`);

      if (!slideXml.includes('<a:hlinkClick')) {
        slideXml = slideXml.replace(
          /(<p:cNvPr\b[^>]*?name="(?:Freeform 28|Freeform \d+|Group \d+|[^"]*Pin[^"]*)"[^>]*?>)/gi,
          `$1<a:hlinkClick r:id="${relId}" tooltip="${escapedUrl}"/>`
        );
      }
    }

    // 2. Update tooltips in slideXml if hlinkClick exists
    if (slideXml.includes('<a:hlinkClick')) {
      slideXml = slideXml.replace(/(<a:hlinkClick[^>]*tooltip=")[^"]*(")/gi, `$1${escapedUrl}$2`);
    }

    return { slideXml, relsXml };
  }

  // adinn-new-template slide 4/5 cloning: exact text swaps, plus per-relationship-id image
  // replacement that also recomputes that shape's <a:fillRect> crop so the new photo covers
  // its existing box without stretching (aspect-fit "cover", not distort-to-fill).
  // images: [{ relId, buffer, ext, boxWidthEMU, boxHeightEMU }]
  // clearImageRelId: relationship id whose blipFill should become empty (map placeholder)
  // placeholderText: { offX, offY, extCx, extCy, text } — new centered textbox for that empty area
  // removeGroupNames: names of top-level <p:grpSp> shapes to drop entirely (e.g.
  // Adinn-Direct-Client-format's Site Information card background when the site has none set).
  // removeShapesAtOffset: plain <p:sp> shapes to drop, matched by their exact <a:off y>/<a:ext cy>
  // (that card's description text is a separate sibling shape, not nested inside the group, and
  // its content/x-position vary per example, so position is the only stable match). Both default
  // to empty so existing callers (adinn-new-template) are unaffected.
  async cloneAdinnSiteSlide(
    templateBaseName,
    {
      textReplacements = [],
      images = [],
      clearImageRelId,
      placeholderText,
      boxWidths = [],
      removeGroupNames = [],
      removeShapesAtOffset = [],
      locationUrl,
    } = {}
  ) {
    const slidePath = `ppt/slides/${templateBaseName}.xml`;
    const relsPath = `ppt/slides/_rels/${templateBaseName}.xml.rels`;

    let slideXml = await this.readText(slidePath);
    let relsXml = await this.readText(relsPath);

    if (locationUrl) {
      const updated = this.updateLocationHyperlink(slideXml, relsXml, locationUrl);
      slideXml = updated.slideXml;
      relsXml = updated.relsXml;
    }

    slideXml = removeRedHighlightShapes(slideXml);

    for (const groupName of removeGroupNames) {
      const groupRe = /<p:grpSp>(?:(?!<\/p:grpSp>)[\s\S])*?<\/p:grpSp>/g;
      slideXml = slideXml.replace(groupRe, (block) => (block.includes(`name="${groupName}"`) ? '' : block));
    }

    for (const { offY, extCy } of removeShapesAtOffset) {
      const shapeRe = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
      const marker = `y="${offY}"/><a:ext cx="`;
      slideXml = slideXml.replace(shapeRe, (block) =>
        block.includes(marker) && block.includes(`cy="${extCy}"/>`) ? '' : block
      );
    }

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

    for (const { relId, buffer, ext, boxWidthEMU, boxHeightEMU, fitContain } of images) {
      if (!buffer) continue;
      const newTarget = await this.addMediaFile(buffer, ext);
      relsXml = this.replaceRelTarget(relsXml, relId, newTarget);

      // <a:fillRect> insets shrink the visible image *within* the shape (leaving gaps) — the
      // correct "cover crop" element is <a:srcRect>, which crops the source image itself before
      // the (gap-free) stretch-to-fill-shape happens. Skipped (zero insets) only when "contain"
      // repositioning below can actually run — findPicXfrm refuses shapes nested inside a
      // <p:grpSp> (their raw off/ext aren't real slide-absolute EMUs) or non-rectangular shapes,
      // so fitPicIntoBox silently no-ops for those; falling back to a real cover-crop there avoids
      // leaving the image both uncropped AND unrepositioned, which would force-stretch/distort it.
      const dims = getImageDimensions(buffer, ext);
      const canReposition = fitContain && dims && findPicXfrm(slideXml, relId) !== null;
      const rect = !canReposition && dims ? computeCoverFillRect(dims.width, dims.height, boxWidthEMU, boxHeightEMU) : { l: 0, t: 0, r: 0, b: 0 };
      const blipStretchRe = new RegExp(`<a:blip r:embed="${relId}"\\/><a:stretch><a:fillRect[^/]*\\/><\\/a:stretch>`);
      slideXml = slideXml.replace(
        blipStretchRe,
        `<a:blip r:embed="${relId}"/><a:srcRect l="${rect.l}" t="${rect.t}" r="${rect.r}" b="${rect.b}"/><a:stretch><a:fillRect/></a:stretch>`
      );

      // "Contain" fit (full image, no crop, no stretch) — opt-in per image, used for the two-column
      // site-photo + location-map layouts. Single full-bleed photos leave this off, keeping their
      // existing cover/stretch look.
      if (canReposition) {
        slideXml = fitPicIntoBox(slideXml, relId, dims);
      }
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

  // adinn-photos-only slide2 template: one two-run caption (location run + a separate, smaller
  // "  WxH" size run) and one full-bleed site photo. Cloned once per selected site; the caption
  // runs and the photo's blip/srcRect are matched by their exact reference-template text/markup,
  // so only those two things change and everything else (fonts, position, logo, etc.) is preserved.
  async clonePhotoOnlySlide(templateBaseName, { locationText, sizeText, image, relId = 'rId3', boxWidthEMU, boxHeightEMU } = {}) {
    const slidePath = `ppt/slides/${templateBaseName}.xml`;
    const relsPath = `ppt/slides/_rels/${templateBaseName}.xml.rels`;

    let slideXml = await this.readText(slidePath);
    let relsXml = await this.readText(relsPath);

    slideXml = removeRedHighlightShapes(slideXml);

    slideXml = slideXml.replace('<a:t>Yanaikkal junction</a:t>', `<a:t>${xmlEscape(locationText)}</a:t>`);
    slideXml = slideXml.replace('<a:t>  20x20</a:t>', `<a:t>${xmlEscape(sizeText)}</a:t>`);

    if (image && image.buffer) {
      const newTarget = await this.addMediaFile(image.buffer, image.ext);
      relsXml = this.replaceRelTarget(relsXml, relId, newTarget);

      const dims = getImageDimensions(image.buffer, image.ext);
      const rect = dims ? computeCoverFillRect(dims.width, dims.height, boxWidthEMU, boxHeightEMU) : { l: 0, t: 0, r: 0, b: 0 };
      const target = `<a:blip r:embed="${relId}" cstate="print"><a:lum/></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch>`;
      const replacement = `<a:blip r:embed="${relId}" cstate="print"><a:lum/></a:blip><a:srcRect l="${rect.l}" t="${rect.t}" r="${rect.r}" b="${rect.b}"/><a:stretch><a:fillRect/></a:stretch>`;
      slideXml = slideXml.replace(target, replacement);
    }

    return this._registerClonedSlide(slideXml, relsXml);
  }

  // adinn-photos-only's slide2 reference design has only ever had one full-bleed photo — no
  // second/map box exists in the template. "With Location" mode adds one: the existing photo is
  // narrowed to `leftBoxWidthEMU` and a new picture (or, with no map available, the same
  // "Insert your map image here" placeholder used elsewhere) is inserted beside it at
  // `rightBoxWidthEMU`. Caption text replacement is identical to clonePhotoOnlySlide.
  async clonePhotoWithMapSlide(
    templateBaseName,
    { locationText, sizeText, image, mapImage, mapLabel, relId = 'rId3', leftBoxWidthEMU, rightBoxWidthEMU, gapEMU = 100000 } = {}
  ) {
    const slidePath = `ppt/slides/${templateBaseName}.xml`;
    const relsPath = `ppt/slides/_rels/${templateBaseName}.xml.rels`;

    let slideXml = await this.readText(slidePath);
    let relsXml = await this.readText(relsPath);

    slideXml = removeRedHighlightShapes(slideXml);

    slideXml = slideXml.replace('<a:t>Yanaikkal junction</a:t>', `<a:t>${xmlEscape(locationText)}</a:t>`);
    slideXml = slideXml.replace('<a:t>  20x20</a:t>', `<a:t>${xmlEscape(sizeText)}</a:t>`);

    // The slide's root <p:grpSpPr> carries its own (unrelated, all-zero) <a:xfrm> earlier in the
    // document than the photo's — matching against the whole slideXml would grab that one
    // instead. The <p:pic>...</p:pic> block is located first, and only its own <a:xfrm> is
    // read/replaced, leaving the root group transform (and everything else) untouched.
    const picBlockRe = /<p:pic>[\s\S]*?<\/p:pic>/;
    const picBlockMatch = slideXml.match(picBlockRe);
    const picXfrmRe = /<a:xfrm><a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="\d+" cy="(\d+)"\/>/;
    let offX = 675481;
    let offY = 551656;
    let cy = 6096000;
    const imageDims = image && image.buffer ? getImageDimensions(image.buffer, image.ext) : null;
    let leftFitted = { offX, offY, extCx: leftBoxWidthEMU, extCy: cy };
    if (picBlockMatch) {
      const picBlock = picBlockMatch[0];
      const xfrmMatch = picBlock.match(picXfrmRe);
      if (xfrmMatch) {
        offX = parseInt(xfrmMatch[1], 10);
        offY = parseInt(xfrmMatch[2], 10);
        cy = parseInt(xfrmMatch[3], 10);
        leftFitted = imageDims
          ? computeContainBox(imageDims.width, imageDims.height, offX, offY, leftBoxWidthEMU, cy)
          : { offX, offY, extCx: leftBoxWidthEMU, extCy: cy };
        const newXfrmTag = `<a:xfrm><a:off x="${leftFitted.offX}" y="${leftFitted.offY}"/><a:ext cx="${leftFitted.extCx}" cy="${leftFitted.extCy}"/>`;
        const updatedPicBlock = picBlock.replace(xfrmMatch[0], newXfrmTag);
        slideXml = slideXml.slice(0, picBlockMatch.index) + updatedPicBlock + slideXml.slice(picBlockMatch.index + picBlock.length);
      }
    }

    if (image && image.buffer) {
      const newTarget = await this.addMediaFile(image.buffer, image.ext);
      relsXml = this.replaceRelTarget(relsXml, relId, newTarget);

      // Full original image, no crop
      const rect = { l: 0, t: 0, r: 0, b: 0 };
      const target = `<a:blip r:embed="${relId}" cstate="print"><a:lum/></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch>`;
      const replacement = `<a:blip r:embed="${relId}" cstate="print"><a:lum/></a:blip><a:srcRect l="${rect.l}" t="${rect.t}" r="${rect.r}" b="${rect.b}"/><a:stretch><a:fillRect/></a:stretch>`;
      slideXml = slideXml.replace(target, replacement);
    }

    // rightOffX is derived from the box's nominal (not fitted) width so the gap to the map box
    // stays fixed regardless of the photo's own aspect ratio/letterboxing.
    const rightOffX = offX + leftBoxWidthEMU + gapEMU;

    if (mapImage && mapImage.buffer) {
      const mapTarget = await this.addMediaFile(mapImage.buffer, mapImage.ext);
      const newRelId = `rIdGen${this._nextRelId++}`;
      relsXml = relsXml.replace(
        '</Relationships>',
        `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${mapTarget}"/></Relationships>`
      );
      const dims = getImageDimensions(mapImage.buffer, mapImage.ext);
      // No crop — the pic is placed below at its "contain"-fitted size (mapFitted), so the full
      // map (route + pins) is what that box shows. Sized/positioned against the photo's own
      // fitted height (leftFitted), not the raw box cy, so the two end up the same height.
      const rect = { l: 0, t: 0, r: 0, b: 0 };
      const mapFitted = dims
        ? computeContainBox(dims.width, dims.height, rightOffX, leftFitted.offY, rightBoxWidthEMU, leftFitted.extCy)
        : { offX: rightOffX, offY: leftFitted.offY, extCx: rightBoxWidthEMU, extCy: leftFitted.extCy };
      const shapeId = 9500 + this._nextSlideIndex;
      const mapPic =
        `<p:pic><p:nvPicPr><p:cNvPr id="${shapeId}" name="Map Picture"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${newRelId}"/><a:srcRect l="${rect.l}" t="${rect.t}" r="${rect.r}" b="${rect.b}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr bwMode="white"><a:xfrm><a:off x="${mapFitted.offX}" y="${mapFitted.offY}"/><a:ext cx="${mapFitted.extCx}" cy="${mapFitted.extCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      slideXml = slideXml.replace('</p:spTree>', `${mapPic}</p:spTree>`);

      if (mapLabel) {
        // Anchored to the fitted map rect (not the raw box) so the label sits over the bottom
        // edge of the actual visible map image, not an empty letterbox gap beside/below it —
        // with a small margin so it floats just above that edge instead of touching it.
        const labelBox = computeMapLabelBox(mapFitted);
        const labelShapeId = 9550 + this._nextSlideIndex;
        const labelSp =
          `<p:sp><p:nvSpPr><p:cNvPr name="Route Info" id="${labelShapeId}"/><p:cNvSpPr txBox="true"/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="${labelBox.offX}" y="${labelBox.offY}"/><a:ext cx="${labelBox.extCx}" cy="${labelBox.extCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
          `<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="85000"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>` +
          `<p:txBody><a:bodyPr anchor="ctr" wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
          `<a:r><a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="C2221E"/></a:solidFill></a:rPr><a:t>${xmlEscape(mapLabel)}</a:t></a:r>` +
          `</a:p></p:txBody></p:sp>`;
        slideXml = slideXml.replace('</p:spTree>', `${labelSp}</p:spTree>`);
      }
    } else {
      const shapeId = 9600 + this._nextSlideIndex;
      const placeholderSp =
        `<p:sp><p:nvSpPr><p:cNvPr name="Map Placeholder" id="${shapeId}"/><p:cNvSpPr txBox="true"/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="${rightOffX}" y="${offY}"/><a:ext cx="${rightBoxWidthEMU}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>` +
        `<p:txBody><a:bodyPr anchor="ctr" wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
        `<a:r><a:rPr lang="en-US" sz="1800"><a:solidFill><a:srgbClr val="666666"/></a:solidFill><a:latin typeface="Times New Roman MT"/></a:rPr><a:t>Insert your map image here</a:t></a:r>` +
        `</a:p></p:txBody></p:sp>`;
      slideXml = slideXml.replace('</p:spTree>', `${placeholderSp}</p:spTree>`);
    }

    return this._registerClonedSlide(slideXml, relsXml);
  }

  // Shared by Jagran-template-one and publicis-ooh-template's site-image slide: a picture
  // placeholder plus a caption shape whose text is split across several runs purely as a
  // spell-check artifact in the reference files (every run shares identical rPr — sz/bold/font).
  // Rather than relying on literal old-value text matching (fragile once the split varies), the
  // whole caption paragraph is collapsed into a single new run that reuses the first run's rPr.
  // `captionAnchorMarker` is the structural tag (unique per template) that identifies which
  // <p:sp> is the caption — Jagran-template-one's caption is a placeholder shape (p:ph type=
  // "body"), publicis-ooh-template's is a plain manually-inserted textbox (p:cNvSpPr txBox="1")
  // — so it defaults to the former to keep existing Jagran-template-one behavior unchanged.
  // `fitContain`: skip the cover-crop below entirely (zero insets) — used by callers that are
  // about to reposition this same pic with a separate fitImageContain() call (full image, no
  // crop, letterboxed instead). Applying both would crop the image AND THEN resize/reposition
  // the shape using the image's original (pre-crop) aspect ratio, stretching the already-cropped
  // remainder to fill a box sized for the uncropped photo — a doubly-wrong result.
  async cloneCaptionPhotoSlide(
    templateBaseName,
    { captionText, image, relId = 'rId3', boxWidthEMU, boxHeightEMU, captionAnchorMarker = '<p:ph type="body"[^>]*/>', fitContain = false } = {}
  ) {
    const slidePath = `ppt/slides/${templateBaseName}.xml`;
    const relsPath = `ppt/slides/_rels/${templateBaseName}.xml.rels`;

    let slideXml = await this.readText(slidePath);
    let relsXml = await this.readText(relsPath);

    slideXml = removeRedHighlightShapes(slideXml);

    const captionShapeRe = new RegExp(
      `(<p:sp>(?:(?!<\\/p:sp>)[\\s\\S])*?${captionAnchorMarker}[\\s\\S]*?<p:txBody>)([\\s\\S]*?)(<\\/p:txBody>[\\s\\S]*?<\\/p:sp>)`
    );
    const shapeMatch = slideXml.match(captionShapeRe);
    if (shapeMatch) {
      const txBodyInner = shapeMatch[2];
      const bodyPr = (txBodyInner.match(/<a:bodyPr\/>|<a:bodyPr[^>]*\/>|<a:bodyPr[^>]*>[\s\S]*?<\/a:bodyPr>/) || [])[0] || '<a:bodyPr/>';
      const lstStyle = (txBodyInner.match(/<a:lstStyle\/>|<a:lstStyle>[\s\S]*?<\/a:lstStyle>/) || [])[0] || '<a:lstStyle/>';
      const pPr = (txBodyInner.match(/<a:pPr[^>]*\/>|<a:pPr[^>]*>[\s\S]*?<\/a:pPr>/) || [])[0] || '';
      let rPr = (txBodyInner.match(/<a:rPr[^>]*\/>|<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/) || [])[0] || '<a:rPr lang="en-US"/>';
      rPr = rPr.replace(/\s(err|smtClean)="1"/g, '');
      const newTxBody = `${bodyPr}${lstStyle}<a:p>${pPr}<a:r>${rPr}<a:t>${xmlEscape(captionText)}</a:t></a:r></a:p>`;
      slideXml = slideXml.replace(shapeMatch[0], `${shapeMatch[1]}${newTxBody}${shapeMatch[3]}`);
    }

    if (image && image.buffer) {
      const newTarget = await this.addMediaFile(image.buffer, image.ext);
      relsXml = this.replaceRelTarget(relsXml, relId, newTarget);

      // Mirrors cloneAdinnSiteSlide's images loop: only skip the crop when the caller's later
      // fitImageContain() call will actually be able to reposition this shape (not grouped, a
      // plain rectangle) — otherwise apply a real cover-crop now, since that follow-up call will
      // find the same shape unrepositionable and become a no-op.
      const dims = getImageDimensions(image.buffer, image.ext);
      const canReposition = fitContain && dims && findPicXfrm(slideXml, relId) !== null;
      const rect = !canReposition && dims ? computeCoverFillRect(dims.width, dims.height, boxWidthEMU, boxHeightEMU) : { l: 0, t: 0, r: 0, b: 0 };

      // The blip element is self-closing in some templates (Jagran-template-one) but has child
      // elements (e.g. a useLocalDpi hint) in others (publicis-ooh-template), so both forms of
      // its closing tag are tried, preferring the explicit "</a:blip>" close when present.
      const blipCloseRe = new RegExp(`<a:blip r:embed="${relId}"[\\s\\S]*?<\\/a:blip>`);
      const blipSelfCloseRe = new RegExp(`<a:blip r:embed="${relId}"[^/]*/>`);
      const blipMatch = slideXml.match(blipCloseRe) || slideXml.match(blipSelfCloseRe);
      if (blipMatch) {
        const afterBlip = slideXml.slice(blipMatch.index + blipMatch[0].length);
        const srcRectMatch = afterBlip.match(/^<a:srcRect[^/]*\/>/);
        const oldLength = blipMatch[0].length + (srcRectMatch ? srcRectMatch[0].length : 0);
        const replacement =
          `<a:blip r:embed="${relId}" cstate="print"/>` +
          `<a:srcRect l="${rect.l}" t="${rect.t}" r="${rect.r}" b="${rect.b}"/>`;
        slideXml = slideXml.slice(0, blipMatch.index) + replacement + slideXml.slice(blipMatch.index + oldLength);
      }
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

  // Adinn-Direct-Client-format's slide1 stores its demo customer name ("HAVELLS") as a single
  // plain run, not wrapped in any of the "Maxi Vision"/"Hospital"/"Proposal" keywords that
  // setCoverFields matches against (that matching was written for adinn-new-template's own demo
  // content) — so it never got swapped for the selected client's name. Replaced directly here,
  // by that exact literal reference-template text, instead.
  async setCoverCustomerNameLiteral(oldText, customerLabel) {
    let slideXml = await this.readText('ppt/slides/slide1.xml');
    const target = `<a:t>${oldText}</a:t>`;
    if (slideXml.includes(target)) {
      slideXml = slideXml.replace(target, `<a:t>${xmlEscape(customerLabel)}</a:t>`);
      this.writeText('ppt/slides/slide1.xml', slideXml);
    }
  }

  // adinn-customized-format's slide1/slide2 are meant to be blank pages the user fills in
  // manually later, but the uploaded reference file has a full-bleed picture (a past client's
  // cover design) baked into each as their background fill — this clears that image, turning
  // the shape into a plain white fill, while leaving every other shape on the slide (borders,
  // decorative frames) untouched.
  async clearBackgroundImage(slidePath, relId) {
    let slideXml = await this.readText(slidePath);
    const re = new RegExp(`<a:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`);
    slideXml = slideXml.replace(re, '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>');
    this.writeText(slidePath, slideXml);
  }

  // Widens (and/or repositions) a named top-level <p:grpSp>'s own outer <a:xfrm> — used by
  // "without location" mode to let one shape (e.g. a photo box) expand into space freed by
  // removing another element on the same slide. Only the group's own off/ext change; everything
  // inside the group (and every other shape on the slide) is untouched.
  async resizeNamedGroup(slidePath, groupName, { newOffX, newWidthEMU } = {}) {
    let slideXml = await this.readText(slidePath);
    const groupRe = /<p:grpSp>(?:(?!<\/p:grpSp>)[\s\S])*?<\/p:grpSp>/g;
    let match;
    while ((match = groupRe.exec(slideXml))) {
      if (!match[0].includes(`name="${groupName}"`)) continue;
      // The group's own xfrm may or may not carry a rot="..." attribute (kept as-is either way).
      const xfrmMatch = match[0].match(/<a:xfrm( rot="-?\d+")?><a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/);
      if (xfrmMatch) {
        const [full, rotAttr, x, y, cx, cy] = xfrmMatch;
        const newXfrm = `<a:xfrm${rotAttr || ''}><a:off x="${newOffX ?? x}" y="${y}"/><a:ext cx="${newWidthEMU ?? cx}" cy="${cy}"/>`;
        const updatedBlock = match[0].replace(full, newXfrm);
        slideXml = slideXml.slice(0, match.index) + updatedBlock + slideXml.slice(match.index + match[0].length);
      }
      break;
    }
    this.writeText(slidePath, slideXml);
  }

  // Draws a plain solid-white rectangle on a cloned slide, used to blank out label/border
  // graphics that are baked into the slide LAYOUT (not the slide itself), so removing the
  // slide's own value shapes still leaves the layout's static "State:"/"City:" boxes visible
  // behind them. Inserted before any real content that should sit on top of the cover.
  async insertWhiteCover(slidePath, { offX, offY, extCx, extCy }) {
    let slideXml = await this.readText(slidePath);
    const rect =
      `<p:sp><p:nvSpPr><p:cNvPr name="Location Mode Cover" id="${9600 + this._nextSlideIndex}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${extCx}" cy="${extCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`;
    slideXml = slideXml.replace('</p:spTree>', `${rect}</p:spTree>`);
    this.writeText(slidePath, slideXml);
  }

  // Standalone, post-clone version of cloneAdinnSiteSlide's removeShapesAtOffset — drops plain
  // <p:sp> shapes matched by their exact <a:off y>/<a:ext cy>, on an already-cloned slide part.
  // Used where a slide's per-site clone method (e.g. cloneSlide) has no removeShapesAtOffset
  // option of its own, so shapes are stripped in a separate pass after cloning instead.
  async removeShapesAtOffsets(slidePath, offsets = []) {
    let slideXml = await this.readText(slidePath);
    for (const { offY, extCy } of offsets) {
      const shapeRe = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
      const marker = `y="${offY}"/><a:ext cx="`;
      slideXml = slideXml.replace(shapeRe, (block) =>
        block.includes(marker) && block.includes(`cy="${extCy}"/>`) ? '' : block
      );
    }
    this.writeText(slidePath, slideXml);
  }

  // Inserts a brand-new <p:pic> (a real image, e.g. a fetched route map) or, with no buffer, the
  // same "Insert your map image here" placeholder text used elsewhere, at an arbitrary position.
  // Used by templates whose reference design never had a second/map box at all, so one has to be
  // added rather than merely shown/hidden.
  async insertImageOrPlaceholder(slidePath, relsPath, { offX, offY, extCx, extCy, buffer, ext, placeholderText = 'Insert your map image here' }) {
    let slideXml = await this.readText(slidePath);
    let relsXml = await this.readText(relsPath);
    let fittedRect = { offX, offY, extCx, extCy };

    if (buffer) {
      const target = await this.addMediaFile(buffer, ext);
      const newRelId = `rIdGen${this._nextRelId++}`;
      relsXml = relsXml.replace(
        '</Relationships>',
        `<Relationship Id="${newRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/></Relationships>`
      );
      const dims = getImageDimensions(buffer, ext);
      // No crop — the pic below is placed at its "contain"-fitted size (fittedRect), so the
      // complete route/pins/labels stay visible instead of being cropped or distort-stretched.
      const rect = { l: 0, t: 0, r: 0, b: 0 };
      fittedRect = dims ? computeContainBox(dims.width, dims.height, offX, offY, extCx, extCy) : fittedRect;
      const shapeId = 9700 + this._nextSlideIndex;
      const pic =
        `<p:pic><p:nvPicPr><p:cNvPr id="${shapeId}" name="Map Picture"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${newRelId}"/><a:srcRect l="${rect.l}" t="${rect.t}" r="${rect.r}" b="${rect.b}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr bwMode="white"><a:xfrm><a:off x="${fittedRect.offX}" y="${fittedRect.offY}"/><a:ext cx="${fittedRect.extCx}" cy="${fittedRect.extCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      slideXml = slideXml.replace('</p:spTree>', `${pic}</p:spTree>`);
    } else {
      const shapeId = 9800 + this._nextSlideIndex;
      const sp =
        `<p:sp><p:nvSpPr><p:cNvPr name="Map Placeholder" id="${shapeId}"/><p:cNvSpPr txBox="true"/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${extCx}" cy="${extCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>` +
        `<p:txBody><a:bodyPr anchor="ctr" wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
        `<a:r><a:rPr lang="en-US" sz="1800"><a:solidFill><a:srgbClr val="666666"/></a:solidFill><a:latin typeface="Times New Roman MT"/></a:rPr><a:t>${xmlEscape(placeholderText)}</a:t></a:r>` +
        `</a:p></p:txBody></p:sp>`;
      slideXml = slideXml.replace('</p:spTree>', `${sp}</p:spTree>`);
    }

    this.writeText(slidePath, slideXml);
    this.writeText(relsPath, relsXml);
    return fittedRect;
  }

  // Resizes/repositions an existing <p:pic>'s own <a:xfrm> so `buffer`'s image displays fully
  // "contain"-fit (no crop, no stretch, centered) within its current box on an already-cloned
  // slide part. Used where the box's *final* size is only known after a separate mutation (e.g.
  // narrowing a photo box to make room for a map) has already run — reading the box directly off
  // the slide, rather than needing that final size threaded back through, keeps the two steps
  // independent. No-op when there's no new image or the box/blip can't be repositioned (nested
  // inside a group, or non-rectangular) — callers that pass `fitContain` through to the earlier
  // clone method (which checks the identical condition on this same not-yet-touched shape) will
  // already have applied a real cover-crop instead in that case, so this has nothing left to do.
  async fitImageContain(slidePath, relId, buffer, ext) {
    if (!buffer) return;
    const dims = getImageDimensions(buffer, ext);
    if (!dims) return;
    let slideXml = await this.readText(slidePath);
    if (!findPicXfrm(slideXml, relId)) return;
    slideXml = fitPicIntoBox(slideXml, relId, dims);
    this.writeText(slidePath, slideXml);
  }

  // Draws a small "19 mins • 6.8 km" style route-info strip floating just above the bottom edge
  // of a map box (real Google route maps only — never shown over the "Insert your map image
  // here" placeholder). No-op when `text` is empty/falsy, so callers can pass it unconditionally.
  async insertMapLabel(slidePath, { offX, offY, extCx, extCy, text }) {
    if (!text) return;
    let slideXml = await this.readText(slidePath);
    const shapeId = 9900 + this._nextSlideIndex;
    const labelBox = computeMapLabelBox({ offX, offY, extCx, extCy });
    const sp =
      `<p:sp><p:nvSpPr><p:cNvPr name="Route Info" id="${shapeId}"/><p:cNvSpPr txBox="true"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${labelBox.offX}" y="${labelBox.offY}"/><a:ext cx="${labelBox.extCx}" cy="${labelBox.extCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="85000"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>` +
      `<p:txBody><a:bodyPr anchor="ctr" wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
      `<a:r><a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="C2221E"/></a:solidFill></a:rPr><a:t>${xmlEscape(text)}</a:t></a:r>` +
      `</a:p></p:txBody></p:sp>`;
    slideXml = slideXml.replace('</p:spTree>', `${sp}</p:spTree>`);
    this.writeText(slidePath, slideXml);
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

module.exports = { PptxTemplate, extOf, computeContainBox, getImageDimensions };
