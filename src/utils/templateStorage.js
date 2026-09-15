const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const TEMPLATE_DIR = path.join(__dirname, '..', '..', 'assets', 'proposal-templates');

function saveTemplateFile(file) {
  fs.mkdirSync(TEMPLATE_DIR, { recursive: true });
  const ext = path.extname(file.originalname || '') || '.pptx';
  const fileName = `template-${Date.now()}-${crypto.randomBytes(4).toString('hex')}${ext}`;
  const filePath = path.join(TEMPLATE_DIR, fileName);
  fs.writeFileSync(filePath, file.buffer);
  return `/assets/proposal-templates/${fileName}`;
}

module.exports = { saveTemplateFile, TEMPLATE_DIR };
