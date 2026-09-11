const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const PptxGenJS = require('pptxgenjs');

const GENERATED_DIR = path.join(__dirname, '..', '..', 'generated');

function ensureGeneratedDir() {
  fs.mkdirSync(GENERATED_DIR, { recursive: true });
  return GENERATED_DIR;
}

async function generateProposalPpt(proposal) {
  const dir = ensureGeneratedDir();
  const filePath = path.join(dir, `${proposal.proposalId}.pptx`);

  const pptx = new PptxGenJS();

  const cover = pptx.addSlide();
  cover.addText(`Proposal ${proposal.proposalId}`, { x: 0.5, y: 1, fontSize: 28, bold: true });
  cover.addText(proposal.client?.name || '', { x: 0.5, y: 1.8, fontSize: 18 });
  cover.addText(`Variant: ${proposal.variant || '-'}`, { x: 0.5, y: 2.4, fontSize: 14 });

  const summary = pptx.addSlide();
  summary.addText('Summary', { x: 0.5, y: 0.3, fontSize: 22, bold: true });
  summary.addText(
    [
      `Total Amount: ₹${proposal.totalAmount ?? '-'}`,
      `GST Amount: ₹${proposal.gstAmount ?? '-'}`,
      `Monthly Amount: ₹${proposal.monthlyAmount ?? '-'}`,
      `Media Count: ${proposal.sites?.length ?? 0}`,
    ].join('\n'),
    { x: 0.5, y: 1, fontSize: 14, lineSpacing: 24 }
  );

  const mediaSlide = pptx.addSlide();
  mediaSlide.addText('Selected Media', { x: 0.5, y: 0.3, fontSize: 22, bold: true });
  const rows = [
    ['Media Name', 'Type', 'City', 'Status'],
    ...(proposal.sites || []).map((s) => [s.mediaName, s.mediaType, s.city, s.mediaStatus]),
  ];
  mediaSlide.addTable(rows, { x: 0.5, y: 1, w: 9, fontSize: 10, autoPage: true });

  await pptx.writeFile({ fileName: filePath });
  return `/generated/${path.basename(filePath)}`;
}

function generateProposalExcel(proposal) {
  const dir = ensureGeneratedDir();
  const filePath = path.join(dir, `${proposal.proposalId}.xlsx`);

  const summarySheet = XLSX.utils.json_to_sheet([
    {
      ProposalId: proposal.proposalId,
      Client: proposal.client?.name || '',
      Variant: proposal.variant || '',
      TotalAmount: proposal.totalAmount || 0,
      GstAmount: proposal.gstAmount || 0,
      MonthlyAmount: proposal.monthlyAmount || 0,
    },
  ]);

  const mediaSheet = XLSX.utils.json_to_sheet(
    (proposal.sites || []).map((s) => ({
      MediaName: s.mediaName,
      MediaType: s.mediaType,
      State: s.state,
      City: s.city,
      Location: s.location,
      Status: s.mediaStatus,
      MonthlyAmount: s.monthlyAmount,
    }))
  );

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
  XLSX.utils.book_append_sheet(workbook, mediaSheet, 'Media');
  XLSX.writeFile(workbook, filePath);

  return `/generated/${path.basename(filePath)}`;
}

module.exports = { generateProposalPpt, generateProposalExcel };
