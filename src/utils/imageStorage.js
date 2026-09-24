const { uploadFile } = require('./storageService');

async function saveMediaImage(file) {
  return uploadFile(file.buffer, file.originalname, file.mimetype, 'outdoor-proposal/mediaImage');
}

module.exports = { saveMediaImage };
