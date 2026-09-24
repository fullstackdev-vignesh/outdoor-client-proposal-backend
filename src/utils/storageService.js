const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function buildFileName(originalName) {
  const ext = path.extname(originalName || '') || '.bin';
  return `${Date.now()}-${crypto.randomBytes(6).toString('hex')}${ext}`;
}

async function putToSpaces(buffer, key, mimeType) {
  const { S3Client, PutObjectCommand } = require('@aws-sdk/client-s3');
  const {
    DO_SPACES_ENDPOINT,
    DO_SPACES_REGION,
    DO_SPACES_KEY,
    DO_SPACES_SECRET,
    DO_SPACES_BUCKET,
    DO_SPACES_CDN_URL,
    DO_SPACES_CDN_BASE,
  } = process.env;

  if (!DO_SPACES_ENDPOINT || !DO_SPACES_KEY || !DO_SPACES_SECRET || !DO_SPACES_BUCKET) {
    throw new Error('DigitalOcean Spaces is not configured. Set DO_SPACES_ENDPOINT, DO_SPACES_REGION, DO_SPACES_KEY, DO_SPACES_SECRET, DO_SPACES_BUCKET.');
  }

  // Ensure key always lives inside 'outdoor-proposal/' folder in Space
  let finalKey = String(key || '').replace(/^\/+/, '');
  if (!finalKey.startsWith('outdoor-proposal/')) {
    finalKey = `outdoor-proposal/${finalKey}`;
  }

  const client = new S3Client({
    endpoint: DO_SPACES_ENDPOINT,
    region: DO_SPACES_REGION || 'us-east-1',
    credentials: { accessKeyId: DO_SPACES_KEY, secretAccessKey: DO_SPACES_SECRET },
  });

  await client.send(
    new PutObjectCommand({
      Bucket: DO_SPACES_BUCKET,
      Key: finalKey,
      Body: buffer,
      ACL: 'public-read',
      ContentType: mimeType || 'application/octet-stream',
    })
  );

  const base = DO_SPACES_CDN_BASE || DO_SPACES_CDN_URL || `${DO_SPACES_ENDPOINT.replace('https://', `https://${DO_SPACES_BUCKET}.`)}`;
  return `${base.replace(/\/$/, '')}/${finalKey}`;
}

async function uploadFile(fileOrBuffer, originalName, mimeType, subFolder = 'outdoor-proposal/mediaImage') {
  const mode = (process.env.MEDIA_IMAGE_STORAGE || process.env.STORAGE_MODE || 'local').toLowerCase();
  const useSpace = mode === 'space' || mode === 'spaces' || mode === 'cloud';

  const buffer = Buffer.isBuffer(fileOrBuffer) ? fileOrBuffer : fileOrBuffer.buffer;
  const fileName = buildFileName(originalName);
  const cleanSubFolder = subFolder ? subFolder.replace(/^\/+/, '') : 'outdoor-proposal/mediaImage';

  if (useSpace) {
    return putToSpaces(buffer, `${cleanSubFolder}/${fileName}`, mimeType);
  } else {
    // Local storage
    const localDir = path.join(__dirname, '..', '..', 'uploads', cleanSubFolder);
    fs.mkdirSync(localDir, { recursive: true });
    const localPath = path.join(localDir, fileName);
    fs.writeFileSync(localPath, buffer);
    return `/uploads/${cleanSubFolder}/${fileName}`;
  }
}

// Always uploads to the configured cloud bucket regardless of MEDIA_IMAGE_STORAGE/STORAGE_MODE.
// Used for system-generated output (e.g. proposal PPTX) that must never be kept in local storage.
async function uploadFileToCloud(buffer, fileName, mimeType, subFolder) {
  return putToSpaces(buffer, `${subFolder}/${fileName}`, mimeType);
}

module.exports = { uploadFile, uploadFileToCloud };
