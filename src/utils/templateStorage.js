const fs = require('fs');
const path = require('path');
const { uploadFile } = require('./storageService');

const REFERENCE_ROOT = path.join(__dirname, '..', '..', 'reference');

async function saveTemplateFile(file) {
  const mime = file.mimetype || (file.originalname?.endsWith('.xlsx') ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  return uploadFile(file.buffer, file.originalname, mime, 'outdoor-proposal/proposal-templates');
}

async function saveExcelTemplateFile(file) {
  if (!file || !/\.(xlsx|xls)$/i.test(file.originalname || '')) {
    throw new Error('Only .xlsx or .xls files are allowed for Excel templates');
  }
  const mime = file.mimetype || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  return uploadFile(file.buffer, file.originalname, mime, 'outdoor-proposal/proposal-templates');
}

function safeFolderName(name) {
  return String(name || '').trim().replace(/[\\/:*?"<>|]/g, '-') || 'template';
}

// Stores PPT master templates at reference/{template name}/{template name}.pptx,
// renaming the uploaded file so the original filename is never persisted.
async function savePptxTemplateFile(file, body) {
  if (!file || !/\.pptx$/i.test(file.originalname || '')) {
    throw new Error('Only .pptx files are allowed for PPT templates');
  }

  const folderName = safeFolderName(body?.name);
  const dir = path.join(REFERENCE_ROOT, folderName);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, `${folderName}.pptx`), file.buffer);

  return `/reference/${folderName}/${folderName}.pptx`;
}

module.exports = { saveTemplateFile, saveExcelTemplateFile, savePptxTemplateFile };
