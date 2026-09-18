const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function buildFileName(originalName) {
  const ext = path.extname(originalName || '') || '.bin';
  return `${Date.now()}-${crypto.randomBytes(6).toString('hex')}${ext}`;
}

async function uploadFile(fileOrBuffer, originalName, mimeType, subFolder = 'media') {
  const mode = (process.env.MEDIA_IMAGE_STORAGE || process.env.STORAGE_MODE || 'local').toLowerCase();
  const useSpace = mode === 'space' || mode === 'spaces' || mode === 'cloud';

  const buffer = Buffer.isBuffer(fileOrBuffer) ? fileOrBuffer : fileOrBuffer.buffer;
  const fileName = buildFileName(originalName);

  if (useSpace) {
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

    const key = `${subFolder}/${fileName}`;
    await client.send(
      new PutObjectCommand({
        Bucket: DO_SPACES_BUCKET,
        Key: key,
        Body: buffer,
        ACL: 'public-read',
        ContentType: mimeType || 'application/octet-stream',
        // Baked into the object's own metadata (not a query param) — DO Spaces then returns
        // this Content-Disposition on every request for the object's public URL, forever, so
        // the browser Save-As name is correct regardless of which frontend code path/cache
        // served the link. Without this, cloud-stored files always download under their
        // randomized physical filename (see buildFileName) since there's no local Express
        // route in front of them to attach headers at request time.
        ContentDisposition: `attachment; filename="${String(originalName || fileName).replace(/["\r\n]/g, '')}"`,
      })
    );

    const base = DO_SPACES_CDN_URL || `${DO_SPACES_ENDPOINT.replace('https://', `https://${DO_SPACES_BUCKET}.`)}`;
    return `${base.replace(/\/$/, '')}/${key}`;
  } else {
    // Local storage
    const localDir = path.join(__dirname, '..', '..', 'uploads', subFolder);
    fs.mkdirSync(localDir, { recursive: true });
    const localPath = path.join(localDir, fileName);
    fs.writeFileSync(localPath, buffer);
    return `/uploads/${subFolder}/${fileName}`;
  }
}

module.exports = { uploadFile };
