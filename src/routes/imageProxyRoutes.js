const express = require('express');
const http = require('http');

// Site images live on a plain-http server (no https). Browsers block/upgrade http images on
// the https frontend (mixed content), so they are streamed through this https backend instead.
// Only allowlisted hosts can be fetched — otherwise this would be an open proxy.
const ALLOWED_HOSTS = (process.env.IMAGE_PROXY_HOSTS || '68.178.205.50')
  .split(',')
  .map((h) => h.trim())
  .filter(Boolean);

const router = express.Router();

router.get('/', (req, res) => {
  let target;
  try {
    target = new URL(req.query.url);
  } catch {
    return res.status(400).json({ message: 'Invalid url' });
  }
  if (target.protocol !== 'http:' || !ALLOWED_HOSTS.includes(target.hostname)) {
    return res.status(403).json({ message: 'Host not allowed' });
  }

  const upstream = http.get(target, { timeout: 15000 }, (imgRes) => {
    const type = imgRes.headers['content-type'] || '';
    if (imgRes.statusCode !== 200 || !type.startsWith('image/')) {
      imgRes.resume();
      return res.status(imgRes.statusCode === 200 ? 415 : 502).end();
    }
    res.set('Content-Type', type);
    if (imgRes.headers['content-length']) res.set('Content-Length', imgRes.headers['content-length']);
    res.set('Cache-Control', 'public, max-age=86400');
    imgRes.pipe(res);
  });
  upstream.on('timeout', () => upstream.destroy(new Error('timeout')));
  upstream.on('error', () => {
    if (!res.headersSent) res.status(502).end();
  });
});

module.exports = router;
