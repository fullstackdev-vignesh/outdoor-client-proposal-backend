const fs = require('fs');
const path = require('path');
const { PptxTemplate, SITE_SPEC_TEMPLATE, SITE_MAP_TEMPLATE, THANK_YOU_SLIDE, extOf } = require('./pptxTemplateEngine');
const { generateExcelFromTemplate } = require('./excelTemplateEngine');
const { getRouteMapBuffer } = require('./mapService');

const GENERATED_DIR = path.join(__dirname, '..', '..', 'generated');
const BACKEND_ROOT = path.join(__dirname, '..', '..');

function ensureGeneratedDir() {
  fs.mkdirSync(GENERATED_DIR, { recursive: true });
  return GENERATED_DIR;
}

function localImage(image) {
  if (!image || /^(https?:|data:|blob:)/.test(image)) return null;
  const abs = path.join(BACKEND_ROOT, image.replace(/^\//, ''));
  if (!fs.existsSync(abs)) return null;
  return { buffer: fs.readFileSync(abs), ext: extOf(abs) || 'jpg' };
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

  const tpl = await PptxTemplate.load();
  const client = proposal.client || {};
  const sites = proposal.sites || [];

  await tpl.setCoverFields({
    customerLabel: customerLabel(proposal),
    dateLabel: formatDisplayDate(new Date()),
  });

  const insertedBaseNames = [];

  for (const site of sites) {
    const siteImage = localImage(site.image);
    const sizeLabel = site.width && site.height ? `${site.width}x${site.height}` : '';
    const specTextReplacements = [
      ['Chennai', site.city || '-'],
      ['Hoarding', site.mediaType || '-'],
      ['Frontlit', site.illumination || '-'],
      ['40x30', sizeLabel || '-'],
      ['1', String(site.quantity || 1)],
      ['Periyanayakanpalayam bridge towards Mettupalayam 40x30', `${site.location || site.areaName || site.city} ${sizeLabel}`.trim()],
    ];

    const specBase = await tpl.cloneSlide(SITE_SPEC_TEMPLATE, {
      textReplacements: specTextReplacements,
      imageReplacements: siteImage ? [{ relId: 'rId2', buffer: siteImage.buffer, ext: siteImage.ext }] : [],
    });
    insertedBaseNames.push(specBase);

    const hasCoords = site.latitude && site.longitude && client.latitude && client.longitude;
    if (hasCoords) {
      const mapBuffer = await getRouteMapBuffer({
        fromLat: client.latitude,
        fromLng: client.longitude,
        toLat: site.latitude,
        toLng: site.longitude,
      });
      if (mapBuffer) {
        const mapBase = await tpl.cloneSlide(SITE_MAP_TEMPLATE, {
          textReplacements: [
            ['Periyanayakanpalayam bridge towards Mettupalayam 40x30', `${site.location || site.areaName || site.city} ${sizeLabel}`.trim()],
          ],
          imageReplacements: [
            ...(siteImage ? [{ relId: 'rId2', buffer: siteImage.buffer, ext: siteImage.ext }, { relId: 'rId6', buffer: siteImage.buffer, ext: siteImage.ext }] : []),
            { relId: 'rId7', buffer: mapBuffer, ext: 'png' },
          ],
        });
        insertedBaseNames.push(mapBase);
      }
    }
  }

  if (insertedBaseNames.length) {
    await tpl.insertSlides(insertedBaseNames, 'slide3');
  }
  await tpl.removeFromSlideOrder(SITE_SPEC_TEMPLATE);
  await tpl.removeFromSlideOrder(SITE_MAP_TEMPLATE);

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
