const { uploadFile } = require('./storageService');

async function saveTemplateFile(file) {
  const mime = file.mimetype || (file.originalname?.endsWith('.xlsx') ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  return uploadFile(file.buffer, file.originalname, mime, 'proposal-templates');
}

module.exports = { saveTemplateFile };
