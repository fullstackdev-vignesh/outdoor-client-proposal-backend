const fs = require('fs');
const path = require('path');
const { PptxTemplate, extOf } = require('./pptxTemplateEngine');
const { generateExcelFromTemplate } = require('./excelTemplateEngine');
const { getRouteMapBuffer } = require('./mapService');
const { uploadFile, uploadFileToCloud } = require('./storageService');
const PPTTemplate = require('../models/PPTTemplate');

const BACKEND_ROOT = path.join(__dirname, '..', '..');

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

async function generateProposalPpt(proposal) {
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
  if (typeof proposal.pptTemplate === 'object' && proposal.pptTemplate.fileUrl) {
    templateFileUrl = proposal.pptTemplate.fileUrl;
  } else if (typeof proposal.pptTemplate === 'object') {
    templateFileUrl = null;
  } else {
    const tplDoc = await PPTTemplate.findById(proposal.pptTemplate);
    if (!tplDoc) {
      throw new Error('Selected PPT template no longer exists');
    }
    templateFileUrl = tplDoc.fileUrl;
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

  // Generated proposal PPTX always lives in the cloud bucket, never on local disk.
  // The folder is keyed off the client name + the proposal's creation date (not "today"),
  // and the file name reuses the existing proposalId convention, so a refresh/regenerate
  // overwrites the same cloud object instead of piling up duplicates.
  const clientNameSafe = sanitizePathSegment(client.name);
  const dateSegment = (proposal.createdAt ? new Date(proposal.createdAt) : new Date()).toISOString().slice(0, 10);
  const folder = `ooh-proposals/${clientNameSafe}-${dateSegment}`;

  let pptUrl;
  try {
    pptUrl = await uploadFileToCloud(
      buffer,
      `${proposal.proposalId}.pptx`,
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      folder
    );
  } catch (uploadErr) {
    throw new Error(`Failed to upload generated PPT to cloud storage: ${uploadErr.message}`);
  }

  return pptUrl;
}

function generateProposalExcel(proposal) {
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

  return generateExcelFromTemplate(rows).then(async (buffer) => {
    let excelUrl;
    try {
      excelUrl = await uploadFile(
        buffer,
        `${proposal.proposalId}.xlsx`,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'generated'
      );
    } catch (spaceErr) {
      console.warn('Cloud storage upload failed for Excel, falling back to local storage:', spaceErr.message);
      const localDir = path.join(BACKEND_ROOT, 'uploads', 'generated');
      fs.mkdirSync(localDir, { recursive: true });
      const localPath = path.join(localDir, `${proposal.proposalId}.xlsx`);
      fs.writeFileSync(localPath, buffer);
      excelUrl = `/uploads/generated/${proposal.proposalId}.xlsx`;
    }
    return excelUrl;
  });
}

module.exports = { generateProposalPpt, generateProposalExcel };
