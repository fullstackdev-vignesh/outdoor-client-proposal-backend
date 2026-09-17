const fs = require('fs');
const path = require('path');
const { PptxTemplate, extOf } = require('./pptxTemplateEngine');
const { generateExcelFromTemplate } = require('./excelTemplateEngine');
const { getExcelConfig } = require('../config/excelTemplateConfigs');
const { getRouteMapBuffer } = require('./mapService');
const { uploadFile } = require('./storageService');
const PPTTemplate = require('../models/PPTTemplate');
const ExcelTemplate = require('../models/ExcelTemplate');

const BACKEND_ROOT = path.join(__dirname, '..', '..');

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
  let templateFileUrl = null;
  if (proposal.pptTemplate) {
    if (typeof proposal.pptTemplate === 'object' && proposal.pptTemplate.fileUrl) {
      templateFileUrl = proposal.pptTemplate.fileUrl;
    } else if (typeof proposal.pptTemplate === 'string') {
      const tplDoc = await PPTTemplate.findById(proposal.pptTemplate);
      if (tplDoc) templateFileUrl = tplDoc.fileUrl;
    }
  }

  const tpl = await PptxTemplate.load(templateFileUrl);
  const client = proposal.client || {};
  const sites = proposal.sites || [];

  await tpl.setCoverFields({
    customerLabel: customerLabel(proposal),
    dateLabel: formatDisplayDate(new Date()),
  });

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

  let pptUrl;
  try {
    pptUrl = await uploadFile(
      buffer,
      `${proposal.proposalId}.pptx`,
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'generated'
    );
  } catch (spaceErr) {
    console.warn('Cloud storage upload failed for PPT, falling back to local storage:', spaceErr.message);
    const localDir = path.join(BACKEND_ROOT, 'uploads', 'generated');
    fs.mkdirSync(localDir, { recursive: true });
    const localPath = path.join(localDir, `${proposal.proposalId}.pptx`);
    fs.writeFileSync(localPath, buffer);
    pptUrl = `/uploads/generated/${proposal.proposalId}.pptx`;
  }

  return pptUrl;
}

function buildExcelRow(site, index, client) {
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
    vendorName: client.vendorName || '',
    vendorCost: client.vendorCost || 0,
  };
}

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

async function loadTemplateBuffer(fileUrl) {
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
  const rows = (proposal.sites || []).map((s, i) => buildExcelRow(s, i, client));

  const { fileUrl, config } = await resolveExcelTemplate(proposal);
  const buffer = await loadTemplateBuffer(fileUrl);

  const outBuffer = await generateExcelFromTemplate(rows, { buffer: buffer || undefined, config });

  let excelUrl;
  try {
    excelUrl = await uploadFile(
      outBuffer,
      `${proposal.proposalId}.xlsx`,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'generated'
    );
  } catch (spaceErr) {
    console.warn('Cloud storage upload failed for Excel, falling back to local storage:', spaceErr.message);
    const localDir = path.join(BACKEND_ROOT, 'uploads', 'generated');
    fs.mkdirSync(localDir, { recursive: true });
    const localPath = path.join(localDir, `${proposal.proposalId}.xlsx`);
    fs.writeFileSync(localPath, outBuffer);
    excelUrl = `/uploads/generated/${proposal.proposalId}.xlsx`;
  }
  return excelUrl;
}

module.exports = { generateProposalPpt, generateProposalExcel };
