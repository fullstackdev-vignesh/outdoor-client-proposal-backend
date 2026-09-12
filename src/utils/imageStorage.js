const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const UPLOAD_DIR = path.join(__dirname, '..', '..', 'uploads', 'media');

function buildFileName(originalName) {
  const ext = path.extname(originalName || '') || '.jpg';
  return `${Date.now()}-${crypto.randomBytes(6).toString('hex')}${ext}`;
}

async function saveLocal(file) {
  fs.mkdirSync(UPLOAD_DIR, { recursive: true });
  const fileName = buildFileName(file.originalname);
  fs.writeFileSync(path.join(UPLOAD_DIR, fileName), file.buffer);
  return `/uploads/media/${fileName}`;
}

async function saveToSpace(file) {
  const { S3Client, PutObjectCommand } = require('@aws-sdk/client-s3');
  const {
    DO_SPACES_ENDPOINT,
    DO_SPACES_REGION,
    DO_SPACES_KEY,
    DO_SPACES_SECRET,
    DO_SPACES_BUCKET,
    DO_SPACES_CDN_URL,
  } = process.env;

  if (!DO_SPACES_ENDPOINT || !DO_SPACES_KEY || !DO_SPACES_SECRET || !DO_SPACES_BUCKET) {
    throw new Error('DigitalOcean Spaces is not configured. Set DO_SPACES_ENDPOINT, DO_SPACES_REGION, DO_SPACES_KEY, DO_SPACES_SECRET, DO_SPACES_BUCKET.');
  }

  const client = new S3Client({
    endpoint: DO_SPACES_ENDPOINT,
    region: DO_SPACES_REGION || 'us-east-1',
    credentials: { accessKeyId: DO_SPACES_KEY, secretAccessKey: DO_SPACES_SECRET },
  });

  const key = `media/${buildFileName(file.originalname)}`;
  await client.send(
    new PutObjectCommand({
      Bucket: DO_SPACES_BUCKET,
      Key: key,
      Body: file.buffer,
      ACL: 'public-read',
      ContentType: file.mimetype,
    })
  );

  const base = DO_SPACES_CDN_URL || `${DO_SPACES_ENDPOINT.replace('https://', `https://${DO_SPACES_BUCKET}.`)}`;
  return `${base.replace(/\/$/, '')}/${key}`;
}

async function saveMediaImage(file) {
  const mode = (process.env.MEDIA_IMAGE_STORAGE || 'local').toLowerCase();
  if (mode === 'space' || mode === 'spaces' || mode === 'cloud') {
    return saveToSpace(file);
  }
  return saveLocal(file);
}

module.exports = { saveMediaImage };
