const fs = require('fs');
const path = require('path');
const { PptxTemplate, extOf } = require('./pptxTemplateEngine');
const { generateExcelFromTemplate } = require('./excelTemplateEngine');
const { getRouteMapBuffer } = require('./mapService');

const GENERATED_DIR = path.join(__dirname, '..', '..', 'generated');
const BACKEND_ROOT = path.join(__dirname, '..', '..');

function ensureGeneratedDir() {
  fs.mkdirSync(GENERATED_DIR, { recursive: true });
  return GENERATED_DIR;
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

async function generateProposalPpt(proposal) {
  const dir = ensureGeneratedDir();
  const filePath = path.join(dir, `${proposal.proposalId}.pptx`);

  const pptTemplateObj = proposal.pptTemplate || {};
  const tpl = await PptxTemplate.load(pptTemplateObj.fileUrl);
  const client = proposal.client || {};
  const sites = proposal.sites || [];

  await tpl.setCoverFields({
    customerLabel: customerLabel(proposal),
    dateLabel: formatDisplayDate(new Date()),
  });

  const slideFiles = await tpl.getSlideFiles();
  const coverTpl = slideFiles[0] || 'slide1';

  // Read text content from each slide to dynamically identify roles
  const slideTexts = {};
  for (const sf of slideFiles) {
    const xml = await tpl.readText(`ppt/slides/${sf}.xml`);
    const texts = [...xml.matchAll(/<a:t>([^<]+)<\/a:t>/g)].map((m) => m[1].trim()).filter(Boolean);
    slideTexts[sf] = texts;
  }

  let cityDividerTpl = null;
  let siteSpecTpl = null;
  let siteMapTpl = null;
  let thankYouTpl = null;

  for (const sf of slideFiles) {
    const txts = slideTexts[sf].join(' ').toLowerCase();
    if (!cityDividerTpl && (sf === 'slide2' || txts.includes('chennai') || txts.includes('madurai') || txts.includes('city'))) {
      cityDividerTpl = sf;
    }
    if (!siteSpecTpl && (txts.includes('duration') || txts.includes('media type') || txts.includes('illumination') || txts.includes('flyover') || txts.includes('bridge') || txts.includes('hoarding') || txts.includes('unipole') || txts.includes('40x30'))) {
      siteSpecTpl = sf;
    }
    if (!siteMapTpl && (txts.includes('route map') || txts.includes('distance') || txts.includes('map'))) {
      siteMapTpl = sf;
    }
    if (txts.includes('thank you') || txts.includes('thanks') || txts.includes('awaiting your approval') || txts.includes('with regards')) {
      thankYouTpl = sf;
    }
  }

  if (!cityDividerTpl && slideFiles.length >= 2) cityDividerTpl = slideFiles[1];
  if (!siteSpecTpl && slideFiles.length >= 3) siteSpecTpl = slideFiles[2];

  // Group requested proposal sites by city
  const sitesByCity = {};
  for (const site of sites) {
    const city = site.city || 'Other';
    if (!sitesByCity[city]) sitesByCity[city] = [];
    sitesByCity[city].push(site);
  }

  const finalOrderedSlides = [coverTpl];

  for (const [city, citySites] of Object.entries(sitesByCity)) {
    const stateName = citySites[0]?.state || client?.state || 'Tamil Nadu';

    // 1. City / State Divider Slide
    if (cityDividerTpl) {
      const cityDividerBase = await tpl.cloneSlide(
        cityDividerTpl,
        {
          textReplacements: [
            ['Chennai', city],
            ['Madurai', city],
            ['Tamil Nadu', stateName],
            ['City:', `City: ${city}`],
            ['State:', `State: ${stateName}`],
          ],
        },
        { city, state: stateName }
      );
      finalOrderedSlides.push(cityDividerBase);
    }

    // 2. Site Spec Slide for each requested site
    for (const site of citySites) {
      const siteImage = await getImageBuffer(site.image || site.mediaImage);
      const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
      const durationLabel = site.bookingInfo?.durationDays ? `${site.bookingInfo.durationDays} Days` : '30 Days';

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
        ['30 Days', durationLabel],
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
      finalOrderedSlides.push(specBase);

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
          finalOrderedSlides.push(mapBase);
        }
      }
    }
  }

  // 3. Append Thank You slide at the very end
  if (thankYouTpl && thankYouTpl !== coverTpl) {
    finalOrderedSlides.push(thankYouTpl);
  }

  // Set the exact final slide list in presentation.xml (removes all dummy template slides)
  await tpl.setFinalSlideOrder(finalOrderedSlides);

  const buffer = await tpl.save();
  fs.writeFileSync(filePath, buffer);
  return `/generated/${path.basename(filePath)}`;
}

function generateProposalExcel(proposal) {
  const dir = ensureGeneratedDir();
  const filePath = path.join(dir, `${proposal.proposalId}.xlsx`);
  const client = proposal.client || {};

  const rows = (proposal.sites || []).map((s, i) => ({
    siNo: i + 1,
    city: s.city,
    media: s.mediaId,
    location: s.location || s.areaName || '',
    qty: s.quantity || 1,
    width: s.width || 0,
    height: s.height || 0,
    type: s.illumination || '',
    displayCostPerMonth: s.monthlyAmount || 0,
    siteStatus: s.mediaStatus ? s.mediaStatus.charAt(0).toUpperCase() + s.mediaStatus.slice(1) : '',
    vendorName: client.vendorName || '',
    vendorCost: client.vendorCost || 0,
  }));

  return generateExcelFromTemplate(rows).then((buffer) => {
    fs.writeFileSync(filePath, buffer);
    return `/generated/${path.basename(filePath)}`;
  });
}

module.exports = { generateProposalPpt, generateProposalExcel };
