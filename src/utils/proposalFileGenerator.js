const fs = require('fs');
const path = require('path');
const { PptxTemplate } = require('./pptxTemplateEngine');
const { generateExcelFromTemplate } = require('./excelTemplateEngine');
const { getRouteMapBuffer } = require('./mapService');
const { uploadFile } = require('./storageService');
const { getImageBuffer } = require('./mediaImage');
const { resolveGenerator } = require('./pptGenerators');
const PPTTemplate = require('../models/PPTTemplate');

const BACKEND_ROOT = path.join(__dirname, '..', '..');

function customerLabel(proposal) {
  const client = proposal.client || {};
  return client.customerType === 'agency' ? `${client.name} (Agency)` : client.name || 'Customer';
}

function formatDisplayDate(date) {
  return new Intl.DateTimeFormat('en-US', { day: '2-digit', month: 'short', year: 'numeric' }).format(date);
}

async function generateProposalPpt(proposal) {
  let templateName = null;
  let templateFileUrl = null;
  if (proposal.pptTemplate) {
    if (typeof proposal.pptTemplate === 'object' && proposal.pptTemplate._id) {
      templateName = proposal.pptTemplate.name;
      templateFileUrl = proposal.pptTemplate.fileUrl;
    } else {
      const tplDoc = await PPTTemplate.findById(proposal.pptTemplate);
      if (tplDoc) {
        templateName = tplDoc.name;
        templateFileUrl = tplDoc.fileUrl;
      }
    }
  }

  const customGenerator = resolveGenerator(templateName);
  const buffer = customGenerator
    ? await customGenerator({ proposal, client: proposal.client || {}, sites: proposal.sites || [] })
    : await generateLegacyPptBuffer(proposal, templateFileUrl);

  return uploadPptBuffer(proposal, buffer);
}

async function generateLegacyPptBuffer(proposal, templateFileUrl) {
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

  return tpl.save();
}

async function uploadPptBuffer(proposal, buffer) {
  try {
    return await uploadFile(
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
    return `/uploads/generated/${proposal.proposalId}.pptx`;
  }
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
