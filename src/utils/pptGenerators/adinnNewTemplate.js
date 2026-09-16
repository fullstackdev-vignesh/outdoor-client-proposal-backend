const { PptxTemplate } = require('../pptxTemplateEngine');
const { getRouteMapBuffer } = require('../mapService');
const { getImageBuffer } = require('../mediaImage');
const { haversineDistanceKm } = require('../haversine');
const { getImageDimensions } = require('./imageDimensions');

// Base template shipped with the backend (a copy of reference/adinn-new-template/
// Adinn-New-Template.pptx). The original file is never modified — PptxTemplate.load()
// only reads it into an in-memory zip; every mutation below happens on that in-memory
// copy, and generate() returns a brand new output buffer.
const TEMPLATE_REL_PATH = 'assets/proposal-templates/adinn-new-template.pptx';

// Slide layout of the source template (verified against its ppt/presentation.xml):
//   slide1 = Cover, slide2 = About Us, slide3 = Why Adinn,
//   slide4 = Site Specification (cloned per site), slide5 = Site Map (cloned per site),
//   slide6 = Thank You.
const COVER_SLIDE = 'slide1';
const STATIC_SLIDES = ['slide2', 'slide3'];
const SPEC_TEMPLATE_SLIDE = 'slide4';
const MAP_TEMPLATE_SLIDE = 'slide5';
const THANKYOU_SLIDE = 'slide6';

// Exact sample text baked into the source template — used as stable text-run
// anchors for replacement instead of fragile pixel/position matching.
const COVER_DATE_RUN = 'ate: Aug 10, 2026'; // the template splits "Date: ..." into a "D" run + this run
const COVER_CLIENT_RUN = 'Maxi Vision Eye Hospital';
const SPEC_TITLE_RUN = 'Periyanayakanpalayam bridge towards Mettupalayam 40x30';
const SPEC_CITY_RUN = 'Chennai';
const SPEC_SIZE_RUN = '40x30';
const SPEC_MEDIA_TYPE_RUN = 'Hoarding';
const SPEC_ILLUMINATION_RUN = 'Frontlit';
const SPEC_UNIT_RUN = '1';

// Bounding box (EMU) for the site photo inserted on the spec slide — the source
// template has no existing photo placeholder there, so we add a new bordered
// picture in the open area left of the "Media Specifications" column (which
// starts at x=12630560 EMU in the template).
const SPEC_PHOTO_BOX = { x: 800000, y: 1550000, w: 10500000, h: 7900000 };
// Bounding box (EMU) of the two existing image shapes on the map slide
// ("Freeform 9" = rId6 = site-photo slot, "Freeform 12" = rId7 = map slot).
const MAP_SHAPE_BOX = { w: 5643663, h: 5618536 };
const MAP_DISTANCE_BADGE_BOX = { x: 12800000, y: 8600000, w: 3800000, h: 620000 };

function xmlEscape(str) {
  return String(str ?? '').replace(/[<>&'"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' }[c]));
}

function formatCoverDate(date) {
  return new Intl.DateTimeFormat('en-US', { month: 'short', day: '2-digit', year: 'numeric' }).format(date);
}

function customerName(proposal, client) {
  return proposal?.client?.name || client?.name || 'Customer';
}

function siteTitle(site) {
  const location = site.location || site.areaName || site.mediaId || '';
  const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
  return [location, sizeLabel].filter(Boolean).join(' ');
}

/** Crop-to-cover srcRect (OOXML thousandths-of-a-percent) so a replacement image
 * fills its box without stretching/distorting — computed from real pixel
 * dimensions instead of guessing. Returns null when no crop is needed/possible. */
function computeCoverSrcRect(imgW, imgH, boxW, boxH) {
  if (!imgW || !imgH || !boxW || !boxH) return null;
  const imgAspect = imgW / imgH;
  const boxAspect = boxW / boxH;
  if (Math.abs(imgAspect - boxAspect) < 0.01) return null;
  if (imgAspect > boxAspect) {
    const keepFraction = boxAspect / imgAspect;
    const cropEachSide = Math.round(((1 - keepFraction) / 2) * 100000);
    return { l: cropEachSide, t: 0, r: cropEachSide, b: 0 };
  }
  const keepFraction = imgAspect / boxAspect;
  const cropEachSide = Math.round(((1 - keepFraction) / 2) * 100000);
  return { l: 0, t: cropEachSide, r: 0, b: cropEachSide };
}

function srcRectXml(srcRect) {
  return srcRect ? `<a:srcRect l="${srcRect.l}" t="${srcRect.t}" r="${srcRect.r}" b="${srcRect.b}"/>` : '';
}

/** Builds a new, editable <p:pic> element (native PPT picture, not a raster
 * slide) for insertion into a slide's <p:spTree>. */
function buildPicXml({ id, name, relId, box, srcRect }) {
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${xmlEscape(name)}"/>` +
    `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${relId}"/>${srcRectXml(srcRect)}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.w}" cy="${box.h}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:ln w="28575"><a:solidFill><a:srgbClr val="1A1A1A"/></a:solidFill></a:ln></p:spPr></p:pic>`
  );
}

function buildPlaceholderXml({ id, box, label }) {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Placeholder"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.w}" cy="${box.h}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EDEDED"/></a:solidFill>` +
    `<a:ln w="12700"><a:solidFill><a:srgbClr val="C9C9C9"/></a:solidFill></a:ln></p:spPr>` +
    `<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r>` +
    `<a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="999999"/></a:solidFill></a:rPr>` +
    `<a:t>${xmlEscape(label)}</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

function buildDistanceBadgeXml({ id, box, label }) {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Distance Badge"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.w}" cy="${box.h}"/></a:xfrm>` +
    `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="C2221E"/></a:solidFill>` +
    `<a:ln><a:noFill/></a:ln></p:spPr>` +
    `<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r>` +
    `<a:rPr lang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>` +
    `<a:t>${xmlEscape(label)}</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

function insertIntoSpTree(slideXml, fragmentXml) {
  return slideXml.replace('</p:spTree>', `${fragmentXml}</p:spTree>`);
}

/** Repoints an existing shape's <a:blipFill> (matched by its current r:embed)
 * onto a freshly-computed cover-crop srcRect, leaving every other shape's
 * design (position/border/fills/text) exactly as authored in the template. */
function replaceBlipFillSrcRect(slideXml, relId, srcRect) {
  const re = new RegExp(`(<a:blipFill><a:blip r:embed="${relId}"/>)[\\s\\S]*?(<\\/a:blipFill>)`);
  return slideXml.replace(re, `$1${srcRectXml(srcRect)}<a:stretch><a:fillRect/></a:stretch>$2`);
}

async function setCoverFields(tpl, { dateLabel, customerLabel }) {
  let slideXml = await tpl.readText(`ppt/slides/${COVER_SLIDE}.xml`);
  slideXml = slideXml.replace(`<a:t>${COVER_DATE_RUN}</a:t>`, `<a:t>ate: ${xmlEscape(dateLabel)}</a:t>`);
  slideXml = slideXml.replace(`<a:t>${COVER_CLIENT_RUN}</a:t>`, `<a:t>${xmlEscape(customerLabel)}</a:t>`);
  tpl.writeText(`ppt/slides/${COVER_SLIDE}.xml`, slideXml);
}

async function buildSpecSlide(tpl, site, siteImage) {
  const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '-';
  const textReplacements = [
    [SPEC_TITLE_RUN, siteTitle(site)],
    [SPEC_CITY_RUN, site.city || '-'],
    [SPEC_SIZE_RUN, sizeLabel],
    [SPEC_MEDIA_TYPE_RUN, site.mediaType || '-'],
    [SPEC_ILLUMINATION_RUN, site.illumination || '-'],
    [SPEC_UNIT_RUN, String(site.quantity || 1)],
  ];
  const baseName = await tpl.cloneSlide(SPEC_TEMPLATE_SLIDE, { textReplacements });

  const slidePath = `ppt/slides/${baseName}.xml`;
  const relsPath = `ppt/slides/_rels/${baseName}.xml.rels`;
  let slideXml = await tpl.readText(slidePath);
  let relsXml = await tpl.readText(relsPath);

  if (siteImage) {
    const target = await tpl.addMediaFile(siteImage.buffer, siteImage.ext);
    const relId = 'rIdSitePhotoGen';
    relsXml = relsXml.replace(
      '</Relationships>',
      `<Relationship Id="${relId}" Target="${target}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/></Relationships>`
    );
    const dims = getImageDimensions(siteImage.buffer, siteImage.ext);
    const srcRect = dims ? computeCoverSrcRect(dims.width, dims.height, SPEC_PHOTO_BOX.w, SPEC_PHOTO_BOX.h) : null;
    const picXml = buildPicXml({ id: 9500, name: 'Site Photo', relId, box: SPEC_PHOTO_BOX, srcRect });
    slideXml = insertIntoSpTree(slideXml, picXml);
  } else {
    slideXml = insertIntoSpTree(slideXml, buildPlaceholderXml({ id: 9500, box: SPEC_PHOTO_BOX, label: 'Site image not available' }));
  }

  tpl.writeText(slidePath, slideXml);
  tpl.writeText(relsPath, relsXml);
  return baseName;
}

async function buildMapSlide(tpl, site, client, siteImage, fetchRouteMap) {
  const baseName = await tpl.cloneSlide(MAP_TEMPLATE_SLIDE, { textReplacements: [[SPEC_TITLE_RUN, siteTitle(site)]] });

  const slidePath = `ppt/slides/${baseName}.xml`;
  const relsPath = `ppt/slides/_rels/${baseName}.xml.rels`;
  let slideXml = await tpl.readText(slidePath);
  let relsXml = await tpl.readText(relsPath);

  if (siteImage) {
    const target = await tpl.addMediaFile(siteImage.buffer, siteImage.ext);
    relsXml = tpl.replaceRelTarget(relsXml, 'rId6', target);
    const dims = getImageDimensions(siteImage.buffer, siteImage.ext);
    const srcRect = dims ? computeCoverSrcRect(dims.width, dims.height, MAP_SHAPE_BOX.w, MAP_SHAPE_BOX.h) : null;
    slideXml = replaceBlipFillSrcRect(slideXml, 'rId6', srcRect);
  }

  const distanceKm = haversineDistanceKm(client?.latitude, client?.longitude, site.latitude, site.longitude);
  if (distanceKm !== null) {
    const mapBuffer = await fetchRouteMap({
      fromLat: client.latitude,
      fromLng: client.longitude,
      toLat: site.latitude,
      toLng: site.longitude,
    });
    if (mapBuffer) {
      const target = await tpl.addMediaFile(mapBuffer, 'png');
      relsXml = tpl.replaceRelTarget(relsXml, 'rId7', target);
      const dims = getImageDimensions(mapBuffer, 'png') || { width: 640, height: 480 };
      const srcRect = computeCoverSrcRect(dims.width, dims.height, MAP_SHAPE_BOX.w, MAP_SHAPE_BOX.h);
      slideXml = replaceBlipFillSrcRect(slideXml, 'rId7', srcRect);
    }
  }

  const distanceLabel = distanceKm !== null ? `Distance: ${distanceKm.toFixed(1)} km` : 'Distance: N/A';
  slideXml = insertIntoSpTree(slideXml, buildDistanceBadgeXml({ id: 9600, box: MAP_DISTANCE_BADGE_BOX, label: distanceLabel }));

  tpl.writeText(slidePath, slideXml);
  tpl.writeText(relsPath, relsXml);
  return baseName;
}

/**
 * @param {object} args
 * @param {object} [args.deps] - Optional collaborator overrides (used by tests);
 *   production callers omit this and get the real map/image fetchers.
 */
async function generate({ proposal, client, sites, deps = {} }) {
  const fetchRouteMap = deps.getRouteMapBuffer || getRouteMapBuffer;
  const fetchSiteImage = deps.getImageBuffer || getImageBuffer;

  const tpl = await PptxTemplate.load(TEMPLATE_REL_PATH);

  await setCoverFields(tpl, {
    dateLabel: formatCoverDate(new Date()),
    customerLabel: customerName(proposal, client),
  });

  const orderedBaseNames = [COVER_SLIDE, ...STATIC_SLIDES];

  // 2 slides per selected site, in selection order — duplicates of the
  // original Site Specification + Site Map template slides.
  for (const site of sites) {
    const siteImage = await fetchSiteImage(site.mediaImage);
    orderedBaseNames.push(await buildSpecSlide(tpl, site, siteImage));
    orderedBaseNames.push(await buildMapSlide(tpl, site, client, siteImage, fetchRouteMap));
  }

  orderedBaseNames.push(THANKYOU_SLIDE);

  await tpl.setFinalSlideOrder(orderedBaseNames);
  return tpl.save();
}

module.exports = { generate };
