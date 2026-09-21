const fs = require('fs');
const path = require('path');
const { PptxTemplate, extOf } = require('./pptxTemplateEngine');
const { generateExcelFromTemplate } = require('./excelTemplateEngine');
const { getExcelConfig } = require('../config/excelTemplateConfigs');
const { getRouteMapBuffer } = require('./mapService');
const { uploadFileToCloud } = require('./storageService');
const PPTTemplate = require('../models/PPTTemplate');
const ExcelTemplate = require('../models/ExcelTemplate');

const BACKEND_ROOT = path.join(__dirname, '..', '..');

// adinn-new-template slide 4's Media Specifications label/value textboxes are narrower than
// their text at 25pt, causing PowerPoint to wrap single words character-by-character. Widening
// them (position/height/style untouched) keeps each field on one line; the gap to the next
// column (or the slide edge) leaves enough margin that this can never overlap anything.
const ADINN_MEDIA_SPEC_BOX_WIDTHS = [
  { offX: 12630560, offY: 2842286, widthEMU: 2500000 }, // City: (label)
  { offX: 12639846, offY: 4233166, widthEMU: 2500000 }, // Size: (label)
  { offX: 12639846, offY: 5626991, widthEMU: 2500000 }, // Media type: (label)
  { offX: 12635198, offY: 7020816, widthEMU: 2500000 }, // Illumination: (label)
  { offX: 12635198, offY: 8414641, widthEMU: 2500000 }, // Unit: (label)
  { offX: 15377227, offY: 2842286, widthEMU: 2700000 }, // City value
  { offX: 15377227, offY: 4233166, widthEMU: 2700000 }, // Size value
  { offX: 15377227, offY: 5626991, widthEMU: 2700000 }, // Media type value
  { offX: 15377227, offY: 7020816, widthEMU: 2700000 }, // Illumination value
  { offX: 15377227, offY: 8408521, widthEMU: 2700000 }, // Unit value
];

function sanitizePathSegment(value) {
  return (
    String(value || '')
      .trim()
      .replace(/[^a-zA-Z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '') || 'client'
  );
}

async function getImageBuffer(image) {
  if (!image) return null;
  if (!/^https?:\/\//i.test(image)) {
    const abs = path.join(BACKEND_ROOT, image.replace(/^\//, ''));
    if (fs.existsSync(abs)) {
      return { buffer: fs.readFileSync(abs), ext: extOf(abs) || 'jpg' };
    }
    return null;
  }
  try {
    const response = await fetch(image);
    if (!response.ok) return null;
    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const extMatch = image.match(/\.(jpg|jpeg|png|webp|gif)(?:\?|$)/i);
    const ext = extMatch ? extMatch[1].toLowerCase() : 'jpg';
    return { buffer, ext };
  } catch (err) {
    console.error('Failed to download media image:', image, err.message);
    return null;
  }
}

function customerLabel(proposal) {
  const client = proposal.client || {};
  return client.customerType === 'agency' ? `${client.name} (Agency)` : client.name || 'Customer';
}

function formatDisplayDate(date) {
  return new Intl.DateTimeFormat('en-US', { day: '2-digit', month: 'short', year: 'numeric' }).format(date);
}

// Shared generated-file name convention across every template AND both PPT/Excel: "<Client>-
// <DD>-<Month>-<YYYY>-<proposalId>.<ext>" (e.g. "ROTN-19-September-2026-PR-MU83CUXH.pptx"),
// using today's date at generation time — not the proposal's original creation date (that's only
// used for the cloud storage folder segment, so a regenerate overwrites the same folder/object).
function buildGeneratedFileName(proposal, client, now, ext) {
  const clientNameSafe = sanitizePathSegment(client.name);
  const dd = String(now.getDate()).padStart(2, '0');
  const monthWords = now.toLocaleString('en-US', { month: 'long' });
  const yyyy = now.getFullYear();
  return `${clientNameSafe}-${dd}-${monthWords}-${yyyy}-${proposal.proposalId}.${ext}`;
}

function buildGeneratedPptFileName(proposal, client, now) {
  return buildGeneratedFileName(proposal, client, now, 'pptx');
}

async function generateProposalPpt(proposal, { locationMode = 'with' } = {}) {
  if (!proposal.client) {
    throw new Error('Selected client could not be found for this proposal');
  }
  if (!Array.isArray(proposal.sites) || proposal.sites.length === 0) {
    throw new Error('No site is selected for this proposal');
  }
  if (!proposal.pptTemplate) {
    throw new Error('No PPT template is selected for this proposal');
  }

  let templateFileUrl = null;
  let templateName = null;
  if (typeof proposal.pptTemplate === 'object' && proposal.pptTemplate.fileUrl) {
    templateFileUrl = proposal.pptTemplate.fileUrl;
    templateName = proposal.pptTemplate.name;
  } else if (typeof proposal.pptTemplate === 'object') {
    templateFileUrl = null;
    templateName = proposal.pptTemplate.name;
  } else {
    const tplDoc = await PPTTemplate.findById(proposal.pptTemplate);
    if (!tplDoc) {
      throw new Error('Selected PPT template no longer exists');
    }
    templateFileUrl = tplDoc.fileUrl;
    templateName = tplDoc.name;
  }

  if (!templateFileUrl) {
    throw new Error('The selected PPT template has no uploaded PPTX file');
  }
  if (!/^https?:\/\//i.test(templateFileUrl)) {
    const absTemplatePath = path.resolve(BACKEND_ROOT, templateFileUrl.replace(/^\//, ''));
    if (!fs.existsSync(absTemplatePath)) {
      throw new Error('The PPTX file for the selected template could not be found on the server');
    }
  }

  const tpl = await PptxTemplate.load(templateFileUrl);
  const client = proposal.client || {};
  const sites = proposal.sites || [];
  const now = new Date();
  // "without" hides location text/title/map for every template below and, where the layout
  // allows it, lets the site photo expand into the freed space; "with" (the default, for
  // backward compatibility with any existing caller that doesn't pass locationMode) is the
  // original, unchanged behavior.
  const withLocation = locationMode !== 'without';
  const isAdinnNewTemplate = templateName === 'adinn-new-template';
  const isAdinnPhotosOnly = templateName === 'adinn-photos-only';
  const isJagranTemplateTwo = templateName === 'Jagran-template-two';
  const isJagranTemplateOne = templateName === 'Jagran-template-one';
  const isPublicisOohTemplate = templateName === 'publicis-ooh-template';
  const isAdinnDirectClientFormat = templateName === 'Adinn-Direct-Client-format';
  const isAdinnCustomizedFormat = templateName === 'adinn-customized-format';

  await tpl.setCoverFields({
    customerLabel: customerLabel(proposal),
    dateLabel: formatDisplayDate(now),
  });

  if (isAdinnCustomizedFormat) {
    // adinn-customized-format: reference slide1/slide2 are meant to be blank pages the user
    // fills in manually later, but the uploaded file has a past client's full-slide cover image
    // baked into each as its background — cleared to plain white here so every generated
    // proposal starts genuinely blank, regardless of what was in the source file. slide3 is the
    // site-detail template (two image boxes — rId2 large, rId3 smaller — both showing the same
    // site photo, plus a single-run title combining location and size), cloned once per selected
    // site using the same cloneAdinnSiteSlide image-swap logic as adinn-new-template. slide4 in
    // the reference file is just a second filled-in example of the identical layout (not a
    // distinct role, unlike adinn-new-template's spec/map pairing), so — like the extra demo
    // slides ignored in other templates — it's never referenced in the final slide order. slide5
    // ("Thank You") is the real closing slide, kept verbatim at the end. The right-hand box
    // (rId3) shows a real route map (client -> site), fetched the same way as adinn-new-template's
    // site+map slide, instead of duplicating the site photo; when either point lacks coordinates
    // it falls back to the same "Insert your map image here" placeholder adinn-new-template uses.
    await tpl.clearBackgroundImage('ppt/slides/slide1.xml', 'rId2');
    await tpl.clearBackgroundImage('ppt/slides/slide2.xml', 'rId2');

    const slideFiles = await tpl.getSlideFiles();
    const staticFirstSlides = ['slide1', 'slide2'];
    const siteDetailTpl = 'slide3';
    const staticLastSlide = slideFiles[slideFiles.length - 1];

    // "Without Location": no title, no map/right box at all — the left photo box widens to fill
    // the whole area (748675 to 17514869, i.e. Group 2's + Group 4's combined original span),
    // applied once to the template slide before cloning so every per-site clone inherits it.
    if (!withLocation) {
      await tpl.resizeNamedGroup('ppt/slides/slide3.xml', 'Group 2', { newWidthEMU: 16766194 });
    }

    const siteSlideBaseNames = [];
    for (const site of sites) {
      const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
      const locationText = site.location || site.areaName || site.mediaName || site.city || '';
      // Title always shows location+size in both modes — matching adinn-new-template, only the
      // map (present here, replaced by a widened photo below) changes with locationMode.
      const titleText = sizeLabel ? `${locationText} – ${sizeLabel}` : locationText;
      const siteImage = await getImageBuffer(site.mediaImage);

      let mapImage = null;
      if (withLocation) {
        const hasCoords = site.latitude && site.longitude && client.latitude && client.longitude;
        if (hasCoords) {
          const mapBuffer = await getRouteMapBuffer({
            fromLat: client.latitude,
            fromLng: client.longitude,
            toLat: site.latitude,
            toLng: site.longitude,
          });
          if (mapBuffer) mapImage = { buffer: mapBuffer, ext: 'png' };
        }
      }

      const images = [];
      if (siteImage) {
        images.push({
          relId: 'rId2',
          ...siteImage,
          boxWidthEMU: withLocation ? 10576310 : 16766194,
          boxHeightEMU: 7076491,
        });
      }
      if (mapImage) images.push({ relId: 'rId3', ...mapImage, boxWidthEMU: 5775471, boxHeightEMU: 7076491 });

      const slideBase = await tpl.cloneAdinnSiteSlide(siteDetailTpl, {
        textReplacements: [['OMR Padur Nr. Hindustan College twds Solinganallur – 30x25', titleText]],
        images,
        // "Without Location": Group 4 (the rId3 box) must be removed outright, not just left
        // unfilled — it sits in document order after the now-widened Group 2, so leaving it in
        // place would render its own reference-file image on top of the enlarged photo. Group 18
        // and Group 22 are the two decorative map-pin icons positioned inside the map box's area
        // (x=13736435/14567387) — once the photo widens to cover that same area they'd float on
        // top of the site photo, so they're removed alongside Group 4 in this mode.
        removeGroupNames: withLocation ? [] : ['Group 4', 'Group 18', 'Group 22'],
        clearImageRelId: withLocation && !mapImage ? 'rId3' : undefined,
        placeholderText:
          withLocation && !mapImage
            ? {
                offX: 11739398,
                offY: 2105609,
                extCx: 5775471,
                extCy: 7076491,
                text: 'Insert your map image here',
              }
            : undefined,
      });
      siteSlideBaseNames.push(slideBase);
    }

    await tpl.setFinalSlideOrder([...staticFirstSlides, ...siteSlideBaseNames, staticLastSlide]);

    const buffer = await tpl.save();
    const clientNameSafe = sanitizePathSegment(client.name);
    const dateSegment = (proposal.createdAt ? new Date(proposal.createdAt) : now).toISOString().slice(0, 10);
    const folder = `ooh-proposals/${clientNameSafe}-${dateSegment}`;
    const fileName = buildGeneratedPptFileName(proposal, client, now);

    try {
      return await uploadFileToCloud(
        buffer,
        fileName,
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        folder
      );
    } catch (uploadErr) {
      throw new Error(`Failed to upload generated PPT to cloud storage: ${uploadErr.message}`);
    }
  }

  if (isAdinnDirectClientFormat) {
    // Adinn-Direct-Client-format: same visual family and generation approach as
    // adinn-new-template (cover/about/why-us kept as-is, slide4 cloned once per site), but its
    // site-spec slide has no Illumination/Unit fields and no separate site+map slide — instead
    // it has one extra "Site Information" card (a gradient rounded-rect + description text)
    // reused from the SiteInfo master data optionally linked on the site. When a site has no
    // siteInfoId, the whole card group is removed so the spec area stays clean (no empty box).
    await tpl.setCoverDateLabel(formatDisplayDate(now));
    await tpl.setCoverCustomerNameLiteral('HAVELLS', customerLabel(proposal));
    await tpl.resizeTextBox('ppt/slides/slide2.xml', {
      offX: 7609014,
      offY: 2436416,
      newOffX: 7500000,
      newWidthEMU: 3150000,
    });

    const slideFiles = await tpl.getSlideFiles();
    // slide5/slide6 in the reference file are two more filled-in example site slides (not a
    // distinct template role) — like Jagran-template-two's unused middle demo slides, they're
    // excluded here so only the real trailing content (slide7, "Thank You") survives into
    // `remainingSlides` and gets appended after the per-site clones.
    const remainingSlides = slideFiles.filter((f) => !['slide1', 'slide2', 'slide3', 'slide4', 'slide5', 'slide6'].includes(f));

    // Same label/value textbox widths as adinn-new-template's City/Size/Media type rows —
    // identical offsets in this template — minus the Illumination/Unit rows it doesn't have.
    const DIRECT_CLIENT_BOX_WIDTHS = ADINN_MEDIA_SPEC_BOX_WIDTHS.filter(
      (b) => ![7020816, 8414641, 8408521].includes(b.offY)
    );

    // This template's reference design never had a map box at all — "With Location" adds one
    // by narrowing the photo group (Group 20) and inserting a real route map (or the usual
    // "Insert your map image here" placeholder) beside it, same idea as adinn-photos-only. Since
    // the whole Media Specifications side of the slide is now hidden in this mode, the map box
    // widens to use that freed space too, so the two boxes span edge-to-edge like
    // adinn-new-template's own site+map slide (slide width 18288000, small right margin only).
    // Photo box width, map box position/width match adinn-new-template's own slide5 (site+map)
    // values exactly, per the user's explicit request — only PHOTO_BOX_OFF_Y/PHOTO_BOX_HEIGHT
    // stay Direct-Client-format's own (that template's vertical geometry differs slightly).
    const PHOTO_BOX_OFF_Y = 1469733;
    const PHOTO_BOX_HEIGHT = 7736815;
    const PHOTO_BOX_LEFT_WIDTH = 10484172;
    const PHOTO_BOX_RIGHT_X = 11605227;
    const PHOTO_BOX_RIGHT_WIDTH = 6508010;
    if (withLocation) {
      await tpl.resizeNamedGroup('ppt/slides/slide4.xml', 'Group 20', { newWidthEMU: PHOTO_BOX_LEFT_WIDTH });
    }

    const siteSlideBaseNames = [];
    for (const site of sites) {
      const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
      // "With Location" matches adinn-new-template's own title convention exactly
      // (location + size together); "Without Location" keeps location only, since Size is
      // already shown separately in the Media Specifications panel that stays visible there.
      const locationText = site.location || site.areaName || site.mediaName || site.city || '';
      const titleText = withLocation ? `${locationText} ${sizeLabel}`.trim() : locationText;
      const siteImage = await getImageBuffer(site.mediaImage);
      const siteInfo = site.siteInfoId && typeof site.siteInfoId === 'object' ? site.siteInfoId : null;

      const textReplacements = [
        ['Little Mount (Chinnamalai) Towards sardar patel rd, adyar/FL', titleText],
        ['Chennai', site.city || '-'],
        ['25x25', sizeLabel || '-'],
        ['Unipole', site.mediaType || '-'],
      ];
      if (siteInfo?.description) {
        textReplacements.push([
          'This site is strategically important due to heavy daily traffic and strong visibility from multiple approach directions, ensuring high audience exposure.',
          siteInfo.description,
        ]);
      }

      const slide4Base = await tpl.cloneAdinnSiteSlide('slide4', {
        textReplacements,
        images: siteImage
          ? [{ relId: 'rId6', ...siteImage, boxWidthEMU: withLocation ? PHOTO_BOX_LEFT_WIDTH : 11366193, boxHeightEMU: PHOTO_BOX_HEIGHT }]
          : [],
        boxWidths: withLocation ? [] : DIRECT_CLIENT_BOX_WIDTHS,
        // "With Location" removes the Site Info card outright (not just when the site has none),
        // since the whole Media Specifications side of the slide is hidden in that mode.
        removeGroupNames: withLocation || !siteInfo?.description ? ['Group 29'] : [],
        // TextBox 18 ("Who we are") is a leftover decorative fragment normally hidden entirely
        // behind the full-width photo — shrinking the photo for "With Location" exposes it,
        // overlapping the new map box, so it's removed only in that split layout. In that same
        // mode the whole Media Specifications panel (heading + City/Size/Media type label+value
        // pairs, each pair sharing one offY/cy) is removed too, leaving only image+map+title.
        removeShapesAtOffset: [
          ...(withLocation || !siteInfo?.description ? [{ offY: 7117158, extCy: 1416150 }] : []),
          ...(withLocation
            ? [
                { offY: 2436416, extCy: 558271 }, // "Who we are"
                { offY: 1611727, extCy: 581025 }, // "Media Specifications" heading
                { offY: 2842286, extCy: 469900 }, // City label + value
                { offY: 4233166, extCy: 469900 }, // Size label + value
                { offY: 5626991, extCy: 469900 }, // Media type label + value
              ]
            : []),
        ],
      });

      if (withLocation) {
        let mapImage = null;
        const hasCoords = site.latitude && site.longitude && client.latitude && client.longitude;
        if (hasCoords) {
          const mapBuffer = await getRouteMapBuffer({
            fromLat: client.latitude,
            fromLng: client.longitude,
            toLat: site.latitude,
            toLng: site.longitude,
          });
          if (mapBuffer) mapImage = { buffer: mapBuffer, ext: 'png' };
        }
        await tpl.insertImageOrPlaceholder(`ppt/slides/${slide4Base}.xml`, `ppt/slides/_rels/${slide4Base}.xml.rels`, {
          offX: PHOTO_BOX_RIGHT_X,
          offY: PHOTO_BOX_OFF_Y,
          extCx: PHOTO_BOX_RIGHT_WIDTH,
          extCy: PHOTO_BOX_HEIGHT,
          buffer: mapImage?.buffer,
          ext: mapImage?.ext,
        });
      }

      siteSlideBaseNames.push(slide4Base);
    }

    await tpl.setFinalSlideOrder(['slide1', 'slide2', 'slide3', ...siteSlideBaseNames, ...remainingSlides]);

    const buffer = await tpl.save();
    const clientNameSafe = sanitizePathSegment(client.name);
    const dateSegment = (proposal.createdAt ? new Date(proposal.createdAt) : now).toISOString().slice(0, 10);
    const folder = `ooh-proposals/${clientNameSafe}-${dateSegment}`;
    const fileName = buildGeneratedPptFileName(proposal, client, now);

    try {
      return await uploadFileToCloud(
        buffer,
        fileName,
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        folder
      );
    } catch (uploadErr) {
      throw new Error(`Failed to upload generated PPT to cloud storage: ${uploadErr.message}`);
    }
  }

  if (isAdinnNewTemplate) {
    // adinn-new-template: slide 1 date/name, slides 2-3 preserved as-is (implemented earlier).
    // Slides 4 (site specification) and 5 (site + map) are cloned once per selected site, in
    // selection order; the final/remaining slides (6+) are left exactly as in the reference
    // file and appended after all the per-site pairs (their own rules come in a later change).
    await tpl.setCoverDateLabel(formatDisplayDate(now));

    // Slide 2's "Who we are" heading wraps to 2 lines at its original width; widen it using the
    // small gaps already free on either side (a decorative shape sits just left of it, and the
    // "ABOUT US" card sits just right of it), without moving/resizing anything else on the slide.
    await tpl.resizeTextBox('ppt/slides/slide2.xml', {
      offX: 7609014,
      offY: 2436416,
      newOffX: 7500000,
      newWidthEMU: 3150000,
    });

    const slideFiles = await tpl.getSlideFiles();
    const remainingSlides = slideFiles.filter((f) => !['slide1', 'slide2', 'slide3', 'slide4', 'slide5'].includes(f));

    // adinn-new-template has two distinct per-site slide roles: slide4 (photo + Media
    // Specifications panel) and slide5 (photo + map). Rather than cloning both per site (the
    // original behavior), locationMode now picks exactly one role per site: "with" → slide5
    // (map), "without" → slide4 (specification panel) — the title always shows location+size
    // on whichever slide is used, in both modes.
    const siteSlideBaseNames = [];
    for (const site of sites) {
      const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
      const titleText = `${site.location || site.areaName || site.mediaName || site.city || ''} ${sizeLabel}`.trim();
      const siteImage = await getImageBuffer(site.mediaImage);
      const titleReplacement = ['Periyanayakanpalayam bridge towards Mettupalayam 40x30', titleText];

      if (!withLocation) {
        // Slide 4 — site specification: title, bordered site-photo box, media spec values.
        // Only the bordered foreground photo box (rId6) gets the site photo — the full-bleed
        // background shape (rId2) stays exactly as in the reference file, since that's what
        // renders as the plain white backdrop behind the Media Specifications panel.
        const slide4Base = await tpl.cloneAdinnSiteSlide('slide4', {
          textReplacements: [
            titleReplacement,
            ['Chennai', site.city || '-'],
            ['40x30', sizeLabel || '-'],
            ['Hoarding', site.mediaType || '-'],
            ['Frontlit', site.illumination || '-'],
            ['1', site.quantity != null ? String(site.quantity) : '-'],
          ],
          images: siteImage ? [{ relId: 'rId6', ...siteImage, boxWidthEMU: 11366193, boxHeightEMU: 7736815 }] : [],
          boxWidths: ADINN_MEDIA_SPEC_BOX_WIDTHS,
        });
        siteSlideBaseNames.push(slide4Base);
        continue;
      }

      // Slide 5 — site + map: title, bordered site-photo box, map image cleared to a placeholder.
      // Same reasoning as slide 4: leave the full-bleed background (rId2) untouched.
      const slide5Base = await tpl.cloneAdinnSiteSlide('slide5', {
        textReplacements: [titleReplacement],
        images: siteImage ? [{ relId: 'rId9', ...siteImage, boxWidthEMU: 10484172, boxHeightEMU: 7646052 }] : [],
        clearImageRelId: 'rId5',
        placeholderText: {
          offX: 11605227,
          offY: 1587800,
          extCx: 6508010,
          extCy: 7646052,
          text: 'Insert your map image here',
        },
      });
      siteSlideBaseNames.push(slide5Base);
    }

    await tpl.setFinalSlideOrder(['slide1', 'slide2', 'slide3', ...siteSlideBaseNames, ...remainingSlides]);

    const buffer = await tpl.save();
    const clientNameSafe = sanitizePathSegment(client.name);
    const dateSegment = (proposal.createdAt ? new Date(proposal.createdAt) : now).toISOString().slice(0, 10);
    const folder = `ooh-proposals/${clientNameSafe}-${dateSegment}`;
    const fileName = buildGeneratedPptFileName(proposal, client, now);

    try {
      return await uploadFileToCloud(
        buffer,
        fileName,
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        folder
      );
    } catch (uploadErr) {
      throw new Error(`Failed to upload generated PPT to cloud storage: ${uploadErr.message}`);
    }
  }

  if (isAdinnPhotosOnly) {
    // adinn-photos-only: reference slide1 is a city-divider template ("Madurai" placeholder +
    // logo), slide2 is a single site-photo template (two-run caption + one full-bleed photo).
    // Both are cloned as many times as needed, in the user's city/site selection order; none of
    // the 12 reference slides are kept verbatim in the output, only their clones.
    const sitesByCity = {};
    const cityOrder = [];
    for (const site of sites) {
      const city = site.city || 'Other';
      if (!sitesByCity[city]) {
        sitesByCity[city] = [];
        cityOrder.push(city);
      }
      sitesByCity[city].push(site);
    }

    const insertedBaseNames = [];
    for (const city of cityOrder) {
      const cityDividerBase = await tpl.cloneSlide('slide1', { textReplacements: [['Madurai', city]] });
      insertedBaseNames.push(cityDividerBase);

      for (const site of sitesByCity[city]) {
        const siteImage = await getImageBuffer(site.mediaImage);
        const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
        // The caption itself is unaffected by locationMode — location+size always shows, same
        // as every other template. What locationMode does control here is the map: this
        // template's reference design only ever had one full-bleed photo box (no map slot at
        // all), so "With Location" adds a second box beside a narrowed photo showing a real
        // route map (or the same "Insert your map image here" placeholder used elsewhere when
        // coordinates/the map fetch aren't available); "Without Location" keeps the original
        // single full-bleed photo, unchanged.
        const locationText = site.location || site.areaName || site.mediaName || site.city || '-';
        const sizeText = sizeLabel ? `  ${sizeLabel}` : '';

        let photoBase;
        if (withLocation) {
          let mapImage = null;
          const hasCoords = site.latitude && site.longitude && client.latitude && client.longitude;
          if (hasCoords) {
            const mapBuffer = await getRouteMapBuffer({
              fromLat: client.latitude,
              fromLng: client.longitude,
              toLat: site.latitude,
              toLng: site.longitude,
            });
            if (mapBuffer) mapImage = { buffer: mapBuffer, ext: 'png' };
          }
          photoBase = await tpl.clonePhotoWithMapSlide('slide2', {
            locationText,
            sizeText,
            image: siteImage,
            mapImage,
            leftBoxWidthEMU: 5674400,
            rightBoxWidthEMU: 3674401,
          });
        } else {
          photoBase = await tpl.clonePhotoOnlySlide('slide2', {
            locationText,
            sizeText,
            image: siteImage,
            boxWidthEMU: 9448801,
            boxHeightEMU: 6096000,
          });
        }
        insertedBaseNames.push(photoBase);
      }
    }

    await tpl.setFinalSlideOrder(insertedBaseNames);

    const buffer = await tpl.save();
    const clientNameSafe = sanitizePathSegment(client.name);
    const dateSegment = (proposal.createdAt ? new Date(proposal.createdAt) : now).toISOString().slice(0, 10);
    const folder = `ooh-proposals/${clientNameSafe}-${dateSegment}`;
    const fileName = buildGeneratedPptFileName(proposal, client, now);

    try {
      return await uploadFileToCloud(
        buffer,
        fileName,
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        folder
      );
    } catch (uploadErr) {
      throw new Error(`Failed to upload generated PPT to cloud storage: ${uploadErr.message}`);
    }
  }

  if (isPublicisOohTemplate) {
    // publicis-ooh-template: reference slide1 (static "New Client Plan – PPT" cover) is kept
    // verbatim. slide2 ("State – Tamil Nadu") is the state-title template, slide3 ("Coimbatore")
    // is the city-title template, and slide4 (photo + multi-run title caption "Trichy Road –
    // Radha Rani Theatre Bus Stop", no size) is the site-image template — cloned once per
    // state/city/site, in selection order. Sites are grouped by state first, then by city within
    // that state (both in first-seen order), mirroring the same city-grouping pattern used by
    // adinn-photos-only/Jagran-template-one/-two, just with an extra state level on top. slide4's
    // caption is a plain textbox (not a placeholder like Jagran-template-one's), so
    // cloneCaptionPhotoSlide is told to anchor on that structural marker instead.
    const slideFiles = await tpl.getSlideFiles();
    const staticFirstSlide = slideFiles[0];
    const staticLastSlide = slideFiles[slideFiles.length - 1];
    const stateTitleTpl = 'slide2';
    const cityTitleTpl = 'slide3';
    const siteDetailTpl = 'slide4';

    const sitesByState = {};
    const stateOrder = [];
    for (const site of sites) {
      const state = site.state || 'Other';
      if (!sitesByState[state]) {
        sitesByState[state] = { cities: {}, cityOrder: [] };
        stateOrder.push(state);
      }
      const city = site.city || 'Other';
      if (!sitesByState[state].cities[city]) {
        sitesByState[state].cities[city] = [];
        sitesByState[state].cityOrder.push(city);
      }
      sitesByState[state].cities[city].push(site);
    }

    const insertedBaseNames = [];
    for (const state of stateOrder) {
      const stateTitleBase = await tpl.cloneSlide(stateTitleTpl, {
        textReplacements: [['State – Tamil Nadu', `State – ${state}`]],
      });
      insertedBaseNames.push(stateTitleBase);

      const { cities, cityOrder } = sitesByState[state];
      for (const city of cityOrder) {
        const cityTitleBase = await tpl.cloneSlide(cityTitleTpl, { textReplacements: [['Coimbatore', city]] });
        insertedBaseNames.push(cityTitleBase);

        for (const site of cities[city]) {
          const siteImage = await getImageBuffer(site.mediaImage);
          const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
          const locationText = site.location || site.areaName || site.mediaName || '-';
          // Caption always shows location+size in both modes, matching adinn-new-template's
          // convention — only the map (added below for "With Location") changes with locationMode.
          const captionText = sizeLabel ? `${locationText} ${sizeLabel}` : locationText;

          // Original photo box (offX=1209675, offY=436563, cy=6005512) spans 9772650 wide with an
          // equal 1209675 margin on both sides (slide width 12192000). "With Location" splits that
          // same span into a narrower left photo + a right map box (adinn-photos-only's pattern),
          // so the pair together still fits exactly where the single photo used to sit.
          const PHOTO_OFF_X = 1209675;
          const PHOTO_OFF_Y = 436563;
          const PHOTO_HEIGHT = 6005512;
          const LEFT_BOX_WIDTH = withLocation ? 5872650 : 9772650;
          const GAP = 100000;
          const RIGHT_BOX_WIDTH = 3800000;

          const siteBase = await tpl.cloneCaptionPhotoSlide(siteDetailTpl, {
            captionText,
            image: siteImage,
            relId: 'rId2',
            boxWidthEMU: LEFT_BOX_WIDTH,
            boxHeightEMU: PHOTO_HEIGHT,
            captionAnchorMarker: '<p:cNvSpPr txBox="1">',
          });

          if (withLocation) {
            await tpl.resizeTextBox(`ppt/slides/${siteBase}.xml`, {
              offX: PHOTO_OFF_X,
              offY: PHOTO_OFF_Y,
              newWidthEMU: LEFT_BOX_WIDTH,
            });

            let mapImage = null;
            const hasCoords = site.latitude && site.longitude && client.latitude && client.longitude;
            if (hasCoords) {
              const mapBuffer = await getRouteMapBuffer({
                fromLat: client.latitude,
                fromLng: client.longitude,
                toLat: site.latitude,
                toLng: site.longitude,
              });
              if (mapBuffer) mapImage = { buffer: mapBuffer, ext: 'png' };
            }

            await tpl.insertImageOrPlaceholder(
              `ppt/slides/${siteBase}.xml`,
              `ppt/slides/_rels/${siteBase}.xml.rels`,
              {
                offX: PHOTO_OFF_X + LEFT_BOX_WIDTH + GAP,
                offY: PHOTO_OFF_Y,
                extCx: RIGHT_BOX_WIDTH,
                extCy: PHOTO_HEIGHT,
                buffer: mapImage?.buffer,
                ext: mapImage?.ext,
              }
            );
          }

          insertedBaseNames.push(siteBase);
        }
      }
    }

    await tpl.setFinalSlideOrder([staticFirstSlide, ...insertedBaseNames, staticLastSlide]);

    const buffer = await tpl.save();
    const clientNameSafe = sanitizePathSegment(client.name);
    const dateSegment = (proposal.createdAt ? new Date(proposal.createdAt) : now).toISOString().slice(0, 10);
    const folder = `ooh-proposals/${clientNameSafe}-${dateSegment}`;
    const fileName = buildGeneratedPptFileName(proposal, client, now);

    try {
      return await uploadFileToCloud(
        buffer,
        fileName,
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        folder
      );
    } catch (uploadErr) {
      throw new Error(`Failed to upload generated PPT to cloud storage: ${uploadErr.message}`);
    }
  }

  if (isJagranTemplateOne) {
    // Jagran-template-one: reference slide1 ("Coimbatore") is the city-divider template and
    // slide2 (photo + multi-run caption "Avinashi Road Goldwins towards Airport -30x30") is the
    // site-image template — both cloned once per city/site, in selection order, exactly like
    // adinn-photos-only/Jagran-template-two. Slides 3-6 in the reference file are extra filled-in
    // example content (not a distinct "Thank You" slide), so — like Jagran-template-two's own
    // unused middle demo slides — they're simply never referenced in the final slide order below
    // and stay as orphaned, ignored parts of the package. If a future master upload adds a real
    // closing slide, wire it in the same way slide1/slide17 are handled for Jagran-template-two.
    const slideFiles = await tpl.getSlideFiles();
    const cityDividerTpl = slideFiles[0];
    const siteDetailTpl = slideFiles[1];

    const sitesByCity = {};
    const cityOrder = [];
    for (const site of sites) {
      const city = site.city || 'Other';
      if (!sitesByCity[city]) {
        sitesByCity[city] = [];
        cityOrder.push(city);
      }
      sitesByCity[city].push(site);
    }

    const insertedBaseNames = [];
    for (const city of cityOrder) {
      const cityDividerBase = await tpl.cloneSlide(cityDividerTpl, { textReplacements: [['Coimbatore', city]] });
      insertedBaseNames.push(cityDividerBase);

      for (const site of sitesByCity[city]) {
        const siteImage = await getImageBuffer(site.mediaImage);
        const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
        const locationText = site.location || site.areaName || site.mediaName || '-';
        // Caption always shows location+size in both modes — only the map (added below for
        // "With Location") changes with locationMode, matching adinn-new-template's convention.
        const captionText = sizeLabel ? `${locationText} -${sizeLabel}` : locationText;

        // Original photo box (offX=571472, offY=928670, cy=4982010) spans 8072494 wide, ending
        // near the slide's right edge (slide width 9144000). "With Location" splits that same
        // span into a narrower left photo + a right map box, so the pair fits exactly where the
        // single photo used to sit — same pattern as publicis-ooh-template/adinn-photos-only.
        const PHOTO_OFF_X = 571472;
        const PHOTO_OFF_Y = 928670;
        const PHOTO_HEIGHT = 4982010;
        const LEFT_BOX_WIDTH = withLocation ? 4772494 : 8072494;
        const GAP = 100000;
        const RIGHT_BOX_WIDTH = 3200000;

        const siteBase = await tpl.cloneCaptionPhotoSlide(siteDetailTpl, {
          captionText,
          image: siteImage,
          boxWidthEMU: LEFT_BOX_WIDTH,
          boxHeightEMU: PHOTO_HEIGHT,
        });

        if (withLocation) {
          await tpl.resizeTextBox(`ppt/slides/${siteBase}.xml`, {
            offX: PHOTO_OFF_X,
            offY: PHOTO_OFF_Y,
            newWidthEMU: LEFT_BOX_WIDTH,
          });

          let mapImage = null;
          const hasCoords = site.latitude && site.longitude && client.latitude && client.longitude;
          if (hasCoords) {
            const mapBuffer = await getRouteMapBuffer({
              fromLat: client.latitude,
              fromLng: client.longitude,
              toLat: site.latitude,
              toLng: site.longitude,
            });
            if (mapBuffer) mapImage = { buffer: mapBuffer, ext: 'png' };
          }

          await tpl.insertImageOrPlaceholder(
            `ppt/slides/${siteBase}.xml`,
            `ppt/slides/_rels/${siteBase}.xml.rels`,
            {
              offX: PHOTO_OFF_X + LEFT_BOX_WIDTH + GAP,
              offY: PHOTO_OFF_Y,
              extCx: RIGHT_BOX_WIDTH,
              extCy: PHOTO_HEIGHT,
              buffer: mapImage?.buffer,
              ext: mapImage?.ext,
            }
          );
        }

        insertedBaseNames.push(siteBase);
      }
    }

    await tpl.setFinalSlideOrder(insertedBaseNames);

    const buffer = await tpl.save();
    const clientNameSafe = sanitizePathSegment(client.name);
    const dateSegment = (proposal.createdAt ? new Date(proposal.createdAt) : now).toISOString().slice(0, 10);
    const folder = `ooh-proposals/${clientNameSafe}-${dateSegment}`;
    const fileName = buildGeneratedPptFileName(proposal, client, now);

    try {
      return await uploadFileToCloud(
        buffer,
        fileName,
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        folder
      );
    } catch (uploadErr) {
      throw new Error(`Failed to upload generated PPT to cloud storage: ${uploadErr.message}`);
    }
  }

  if (isJagranTemplateTwo) {
    // Jagran-template-two: reference slide1 ("Jagran Engage" title) and the last slide
    // ("Thank you") are kept verbatim — no cloning, no text/image mutation. slide2 is the
    // city-divider template (single red city-name textbox); slide4 is used as the site-detail
    // template because its reference Width/Height values (35.00/25.00) are distinct, unlike
    // slide3's duplicate 20.00/20.00 — exact-text replacement can then target each field
    // unambiguously. Both are cloned once per city/site, in the user's selection order, the
    // same way adinn-photos-only clones its city-divider + site-photo pair.
    const slideFiles = await tpl.getSlideFiles();
    const staticFirstSlide = slideFiles[0];
    const staticLastSlide = slideFiles[slideFiles.length - 1];
    const cityDividerTpl = 'slide2';
    const siteDetailTpl = 'slide4';

    const sitesByCity = {};
    const cityOrder = [];
    for (const site of sites) {
      const city = site.city || 'Other';
      if (!sitesByCity[city]) {
        sitesByCity[city] = [];
        cityOrder.push(city);
      }
      sitesByCity[city].push(site);
    }

    const insertedBaseNames = [];
    for (const city of cityOrder) {
      const cityDividerBase = await tpl.cloneSlide(cityDividerTpl, {
        textReplacements: [['Chennai', city]],
        preserveRedShapes: true,
      });
      insertedBaseNames.push(cityDividerBase);

      for (const site of sitesByCity[city]) {
        const siteImage = await getImageBuffer(site.mediaImage);
        const locationText = site.location || site.areaName || site.mediaName || '-';
        const widthText = site.width != null ? Number(site.width).toFixed(2) : '-';
        const heightText = site.height != null ? Number(site.height).toFixed(2) : '-';
        const sizeLabel = site.width && site.height ? `${widthText}x${heightText}` : '';
        // "With Location" matches adinn-new-template's title convention (location+size combined);
        // "Without Location" keeps location only, since Width/Height already show separately in
        // the visible Media Specifications panel there.
        const titleText = withLocation && sizeLabel ? `${locationText} ${sizeLabel}` : locationText;

        const siteBase = await tpl.cloneSlide(siteDetailTpl, {
          textReplacements: [
            ['Porur EB Office towards Porur Signal', titleText],
            ['Tamil Nadu', site.state || '-'],
            ['Chennai', site.city || '-'],
            ['Hoarding', site.mediaType || '-'],
            ['Not Lit', site.illumination || '-'],
            ['35.00', widthText],
            ['25.00', heightText],
          ],
          imageReplacements: siteImage ? [siteImage] : [],
        });

        if (withLocation) {
          // This template's reference design never had a map box — the site photo already sits
          // in the slide's left ~60% (534554 to 10512177 of a 17279938-wide slide), leaving the
          // Media Specifications panel in the right column. That panel's "State:"/"City:"/etc
          // labels and borders are baked into the slide LAYOUT itself (slideLayout1.xml), not
          // the slide — so removing the slide's own value shapes alone still leaves the empty
          // labelled boxes visible underneath. "With Location" therefore also draws a plain
          // white cover over that whole panel area (everything below the title box, which stays
          // visible), then fills the freed space with a real route map (or the usual
          // placeholder), instead of narrowing the photo like the other single-photo templates.
          await tpl.removeShapesAtOffsets(`ppt/slides/${siteBase}.xml`, [
            { offY: 3220502, extCy: 687600 }, // State
            { offY: 4152859, extCy: 650216 }, // City
            { offY: 5001571, extCy: 685800 }, // Media type
            { offY: 5022156, extCy: 617621 }, // Illumination
            { offY: 5869827, extCy: 678927 }, // Width
            { offY: 5880511, extCy: 638897 }, // Height
            { offY: 6817052, extCy: 638897 }, // Duration
          ]);

          // Title box (offY 1299456, cy 1609098) ends at y=2908554 — the cover starts right
          // below it so the title stays visible, and spans to near the slide's right/bottom edges.
          await tpl.insertWhiteCover(`ppt/slides/${siteBase}.xml`, {
            offX: 10897898,
            offY: 2908554,
            extCx: 6182040,
            extCy: 5532209,
          });

          let mapImage = null;
          const hasCoords = site.latitude && site.longitude && client.latitude && client.longitude;
          if (hasCoords) {
            const mapBuffer = await getRouteMapBuffer({
              fromLat: client.latitude,
              fromLng: client.longitude,
              toLat: site.latitude,
              toLng: site.longitude,
            });
            if (mapBuffer) mapImage = { buffer: mapBuffer, ext: 'png' };
          }

          await tpl.insertImageOrPlaceholder(
            `ppt/slides/${siteBase}.xml`,
            `ppt/slides/_rels/${siteBase}.xml.rels`,
            {
              offX: 10997898,
              offY: 3008554,
              extCx: 5982040,
              extCy: 5332209,
              buffer: mapImage?.buffer,
              ext: mapImage?.ext,
            }
          );
        }

        insertedBaseNames.push(siteBase);
      }
    }

    await tpl.setFinalSlideOrder([staticFirstSlide, ...insertedBaseNames, staticLastSlide]);

    const buffer = await tpl.save();
    const clientNameSafe = sanitizePathSegment(client.name);
    const dateSegment = (proposal.createdAt ? new Date(proposal.createdAt) : now).toISOString().slice(0, 10);
    const folder = `ooh-proposals/${clientNameSafe}-${dateSegment}`;
    const fileName = buildGeneratedPptFileName(proposal, client, now);

    try {
      return await uploadFileToCloud(
        buffer,
        fileName,
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        folder
      );
    } catch (uploadErr) {
      throw new Error(`Failed to upload generated PPT to cloud storage: ${uploadErr.message}`);
    }
  }

  const slideFiles = await tpl.getSlideFiles();
  const coverTpl = 'slide1';
  const cityDividerTpl = slideFiles.find((f) => f === 'slide2') || (slideFiles.length >= 2 ? slideFiles[1] : 'slide1');
  const siteSpecTpl = slideFiles.find((f) => f === 'slide3') || (slideFiles.length >= 3 ? slideFiles[2] : cityDividerTpl);
  const siteMapTpl = slideFiles.find((f) => f === 'slide4') || (slideFiles.length >= 4 ? slideFiles[3] : null);

  const keptTemplates = new Set([coverTpl, cityDividerTpl, siteSpecTpl]);
  if (siteMapTpl) keptTemplates.add(siteMapTpl);

  for (const sf of slideFiles) {
    if (!keptTemplates.has(sf)) {
      await tpl.removeFromSlideOrder(sf);
    }
  }

  // Group sites by city
  const sitesByCity = {};
  for (const site of sites) {
    const city = site.city || 'Other';
    if (!sitesByCity[city]) sitesByCity[city] = [];
    sitesByCity[city].push(site);
  }

  const insertedBaseNames = [];

  for (const [city, citySites] of Object.entries(sitesByCity)) {
    // 1. Clone City Divider Slide
    const cityDividerBase = await tpl.cloneSlide(
      cityDividerTpl,
      {
        textReplacements: [
          ['Chennai', city],
          ['Madurai', city],
          ['Tamil Nadu', proposal.client?.state || 'Tamil Nadu'],
        ],
      },
      { city }
    );
    insertedBaseNames.push(cityDividerBase);

    // 2. Clone Site Specs for each site in this city
    for (const site of citySites) {
      const siteImage = await getImageBuffer(site.mediaImage);
      const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
      const specTextReplacements = [
        ['Chennai', site.city || '-'],
        ['Madurai', site.city || '-'],
        ['Tamil Nadu', site.state || '-'],
        ['Hoarding', site.mediaType || '-'],
        ['Unipole', site.mediaType || '-'],
        ['LED Hoarding', site.mediaType || '-'],
        ['Frontlit', site.illumination || '-'],
        ['Front Lit', site.illumination || '-'],
        ['Backlit', site.illumination || '-'],
        ['Back Lit', site.illumination || '-'],
        ['40x30', sizeLabel || '-'],
        ['40x20', sizeLabel || '-'],
        ['50x30', sizeLabel || '-'],
        ['1', String(site.quantity || 1)],
      ];

      const specBase = await tpl.cloneSlide(
        siteSpecTpl,
        {
          textReplacements: specTextReplacements,
          imageReplacements: siteImage ? [siteImage] : [],
        },
        site
      );
      insertedBaseNames.push(specBase);

      const hasCoords = site.latitude && site.longitude && client.latitude && client.longitude;
      if (hasCoords && siteMapTpl) {
        const mapBuffer = await getRouteMapBuffer({
          fromLat: client.latitude,
          fromLng: client.longitude,
          toLat: site.latitude,
          toLng: site.longitude,
        });
        if (mapBuffer) {
          const mapBase = await tpl.cloneSlide(
            siteMapTpl,
            {
              textReplacements: specTextReplacements,
              imageReplacements: [
                ...(siteImage ? [siteImage, siteImage] : []),
                { buffer: mapBuffer, ext: 'png' },
              ],
            },
            site
          );
          insertedBaseNames.push(mapBase);
        }
      }
    }
  }

  if (insertedBaseNames.length) {
    await tpl.setFinalSlideOrder(['slide1', ...insertedBaseNames]);
  }

  const buffer = await tpl.save();

  // Generated proposal PPTX always lives in the cloud bucket, never on local disk.
  // The folder is keyed off the client name + the proposal's creation date (not "today"),
  // and the file name reuses the existing proposalId convention, so a refresh/regenerate
  // overwrites the same cloud object instead of piling up duplicates.
  const clientNameSafe = sanitizePathSegment(client.name);
  const dateSegment = (proposal.createdAt ? new Date(proposal.createdAt) : new Date()).toISOString().slice(0, 10);
  const folder = `ooh-proposals/${clientNameSafe}-${dateSegment}`;
  const fileName = buildGeneratedPptFileName(proposal, client, now);

  let pptUrl;
  try {
    pptUrl = await uploadFileToCloud(
      buffer,
      fileName,
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      folder
    );
  } catch (uploadErr) {
    throw new Error(`Failed to upload generated PPT to cloud storage: ${uploadErr.message}`);
  }

  return pptUrl;
}

// Generic field shape every Excel format's `columns` config reads from — add a new field here
// first if a new Site field needs to appear in some format, then reference it by name in that
// format's config entry.
function buildExcelRow(site, index) {
  return {
    siNo: index + 1,
    state: site.state || '',
    city: site.city || '',
    media: site.mediaId || '',
    location: site.location || site.areaName || '',
    qty: site.quantity || 1,
    width: site.width || 0,
    height: site.height || 0,
    type: site.illumination || '',
    durationDays: site.bookingInfo?.durationDays || '',
    displayCostPerMonth: site.monthlyAmount || 0,
    printingCost: site.printingCost || 0,
    mountingCost: site.mountingCost || 0,
    siteStatus: site.mediaStatus ? site.mediaStatus.charAt(0).toUpperCase() + site.mediaStatus.slice(1) : '',
  };
}

// Resolves the proposal's Excel Template doc (whether already populated or just an ObjectId
// string) to its uploaded file's URL + the format config that describes that file's layout —
// falling back to the bundled 'generic' (Adinn) config/master file when no template is set.
async function resolveExcelTemplate(proposal) {
  let tplDoc = null;
  if (proposal.excelTemplate) {
    if (typeof proposal.excelTemplate === 'object' && proposal.excelTemplate.fileUrl) {
      tplDoc = proposal.excelTemplate;
    } else if (typeof proposal.excelTemplate === 'string') {
      tplDoc = await ExcelTemplate.findById(proposal.excelTemplate);
    }
  }
  return { fileUrl: tplDoc?.fileUrl || null, config: getExcelConfig(tplDoc?.formatKey || 'generic') };
}

async function loadExcelTemplateBuffer(fileUrl) {
  if (!fileUrl) return null;
  if (/^https?:\/\//i.test(fileUrl)) {
    const res = await fetch(fileUrl);
    if (!res.ok) throw new Error(`Failed to fetch Excel template (${res.status})`);
    return Buffer.from(await res.arrayBuffer());
  }
  const abs = path.resolve(BACKEND_ROOT, fileUrl.replace(/^\//, ''));
  return fs.existsSync(abs) ? fs.readFileSync(abs) : null;
}

async function generateProposalExcel(proposal) {
  const client = proposal.client || {};
  const rows = (proposal.sites || []).map((s, i) => buildExcelRow(s, i));

  const { fileUrl, config } = await resolveExcelTemplate(proposal);
  const buffer = await loadExcelTemplateBuffer(fileUrl);

  const outBuffer = await generateExcelFromTemplate(rows, { buffer: buffer || undefined, config, client });
  const fileName = buildGeneratedFileName(proposal, client, new Date(), 'xlsx');

  try {
    // uploadFileToCloud (unlike uploadFile) uses the given fileName as the storage key as-is,
    // instead of discarding it for a randomized one — same call PPT generation already uses, so
    // the friendly "<Client>-<DD>-<Month>-<YYYY>-<proposalId>.xlsx" name survives into the actual
    // download instead of showing a random storage-generated name.
    return await uploadFileToCloud(
      outBuffer,
      fileName,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'generated'
    );
  } catch (spaceErr) {
    console.warn('Cloud storage upload failed for Excel, falling back to local storage:', spaceErr.message);
    const localDir = path.join(BACKEND_ROOT, 'uploads', 'generated');
    fs.mkdirSync(localDir, { recursive: true });
    fs.writeFileSync(path.join(localDir, fileName), outBuffer);
    return `/uploads/generated/${fileName}`;
  }
}

module.exports = { generateProposalPpt, generateProposalExcel };
