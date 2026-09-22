const fs = require('fs');
const path = require('path');
const { PNG } = require('pngjs');

const GOOGLE_DIRECTIONS_URL = 'https://maps.googleapis.com/maps/api/directions/json';

const ASSETS_DIR = path.join(__dirname, '..', '..', 'assets', 'map-pins');

let pinGreenCache = null;
let pinRedCache = null;

function getPinIcons() {
  if (!pinGreenCache) {
    const greenPath = path.join(ASSETS_DIR, 'pin_green.png');
    if (fs.existsSync(greenPath)) {
      pinGreenCache = PNG.sync.read(fs.readFileSync(greenPath));
    }
  }
  if (!pinRedCache) {
    const redPath = path.join(ASSETS_DIR, 'pin_red_adinn.png');
    if (fs.existsSync(redPath)) {
      pinRedCache = PNG.sync.read(fs.readFileSync(redPath));
    }
  }
  return { greenPin: pinGreenCache, redPin: pinRedCache };
}

function lonToX(lon, zoom) {
  return ((lon + 180) / 360) * Math.pow(2, zoom) * 256;
}

function latToY(lat, zoom) {
  const sin = Math.sin((lat * Math.PI) / 180);
  const rad = Math.log((1 + sin) / (1 - sin)) / 2;
  return (0.5 - rad / (2 * Math.PI)) * Math.pow(2, zoom) * 256;
}

function overlayPinImage(canvas, iconPng, targetX, targetY, scale = 0.22) {
  if (!iconPng) return;
  const scaledW = Math.round(iconPng.width * scale);
  const scaledH = Math.round(iconPng.height * scale);

  // Tip of the pin sits at (targetX, targetY)
  const startX = Math.round(targetX - scaledW / 2);
  const startY = Math.round(targetY - scaledH);

  for (let sy = 0; sy < scaledH; sy++) {
    const destY = startY + sy;
    if (destY < 0 || destY >= canvas.height) continue;

    const srcY0 = sy / scale;
    const srcY1 = (sy + 1) / scale;

    for (let sx = 0; sx < scaledW; sx++) {
      const destX = startX + sx;
      if (destX < 0 || destX >= canvas.width) continue;

      const srcX0 = sx / scale;
      const srcX1 = (sx + 1) / scale;

      let rSum = 0, gSum = 0, bSum = 0, aSum = 0, weightSum = 0;

      const minX = Math.floor(srcX0);
      const maxX = Math.min(Math.ceil(srcX1), iconPng.width);
      const minY = Math.floor(srcY0);
      const maxY = Math.min(Math.ceil(srcY1), iconPng.height);

      for (let py = minY; py < maxY; py++) {
        const wy = Math.min(py + 1, srcY1) - Math.max(py, srcY0);
        for (let px = minX; px < maxX; px++) {
          const wx = Math.min(px + 1, srcX1) - Math.max(px, srcX0);
          const w = wx * wy;

          const srcIdx = (iconPng.width * py + px) << 2;
          const a = iconPng.data[srcIdx + 3] / 255;

          rSum += iconPng.data[srcIdx] * a * w;
          gSum += iconPng.data[srcIdx + 1] * a * w;
          bSum += iconPng.data[srcIdx + 2] * a * w;
          aSum += a * w;
          weightSum += w;
        }
      }

      if (weightSum > 0 && aSum > 0) {
        const finalA = aSum / weightSum;
        const finalR = Math.round(rSum / aSum);
        const finalG = Math.round(gSum / aSum);
        const finalB = Math.round(bSum / aSum);

        const destIdx = (canvas.width * destY + destX) << 2;
        const prevA = canvas.data[destIdx + 3] / 255;

        canvas.data[destIdx] = Math.round(finalR * finalA + canvas.data[destIdx] * (1 - finalA));
        canvas.data[destIdx + 1] = Math.round(finalG * finalA + canvas.data[destIdx + 1] * (1 - finalA));
        canvas.data[destIdx + 2] = Math.round(finalB * finalA + canvas.data[destIdx + 2] * (1 - finalA));
        canvas.data[destIdx + 3] = Math.round((finalA + prevA * (1 - finalA)) * 255);
      }
    }
  }
}

async function renderRouteMapImage({ fromLat, fromLng, toLat, toLng, width = 640, height = 480 }) {
  let routePoints = [];
  let distanceText = null;
  let durationText = null;

  // 1. Try OSRM route service
  try {
    const osrmUrl = `http://router.project-osrm.org/route/v1/driving/${fromLng},${fromLat};${toLng},${toLat}?overview=full&geometries=geojson`;
    const osrmRes = await fetch(osrmUrl, { headers: { 'User-Agent': 'OutdoorProposalApp/1.0' } });
    if (osrmRes.ok) {
      const osrmData = await osrmRes.json();
      const route = osrmData.routes?.[0];
      if (route) {
        if (route.distance) {
          distanceText = `${(route.distance / 1000).toFixed(1)} km`;
        }
        if (route.duration) {
          const mins = Math.round(route.duration / 60);
          durationText = `${mins} mins`;
        }
        if (route.geometry?.coordinates) {
          routePoints = route.geometry.coordinates.map((c) => ({ lat: c[1], lng: c[0] }));
        }
      }
    }
  } catch (err) {
    // Ignore routing failure, fallback to direct line
  }

  // 2. Try Google Directions API if OSRM didn't return route geometry and Google API key is present
  const apiKey = process.env.GOOGLE_MAPS_API_KEY;
  if (routePoints.length === 0 && apiKey) {
    try {
      const dirParams = new URLSearchParams({
        origin: `${fromLat},${fromLng}`,
        destination: `${toLat},${toLng}`,
        mode: 'driving',
        key: apiKey,
      });
      const dirRes = await fetch(`${GOOGLE_DIRECTIONS_URL}?${dirParams.toString()}`);
      if (dirRes.ok) {
        const dirData = await dirRes.json();
        const leg = dirData.routes?.[0]?.legs?.[0];
        if (leg) {
          distanceText = leg.distance?.text || distanceText;
          durationText = leg.duration?.text || durationText;
          if (leg.steps) {
            routePoints = [];
            for (const step of leg.steps) {
              if (step.start_location) routePoints.push({ lat: step.start_location.lat, lng: step.start_location.lng });
              if (step.end_location) routePoints.push({ lat: step.end_location.lat, lng: step.end_location.lng });
            }
          }
        }
      }
    } catch {
      // Fall through to direct line
    }
  }

  if (routePoints.length === 0) {
    routePoints = [{ lat: fromLat, lng: fromLng }, { lat: toLat, lng: toLng }];
  }

  const allLats = [...routePoints.map((p) => p.lat), fromLat, toLat];
  const allLngs = [...routePoints.map((p) => p.lng), fromLng, toLng];

  const minLat = Math.min(...allLats);
  const maxLat = Math.max(...allLats);
  const minLng = Math.min(...allLngs);
  const maxLng = Math.max(...allLngs);

  // Zoom calculation with generous padding (200px) and capped max zoom (15) so the full route and
  // pin markers fit comfortably inside the canvas with ample margin, preventing pins from being cut off.
  let zoom = 15;
  for (; zoom >= 2; zoom--) {
    const x1 = lonToX(minLng, zoom);
    const y1 = latToY(maxLat, zoom);
    const x2 = lonToX(maxLng, zoom);
    const y2 = latToY(minLat, zoom);
    if (Math.abs(x2 - x1) <= width - 200 && Math.abs(y2 - y1) <= height - 200) {
      break;
    }
  }

  const centerX = (lonToX(minLng, zoom) + lonToX(maxLng, zoom)) / 2;
  const centerY = (latToY(minLat, zoom) + latToY(maxLat, zoom)) / 2;
  const originX = centerX - width / 2;
  const originY = centerY - height / 2;

  const canvas = new PNG({ width, height });
  canvas.data.fill(255);

  const minTileX = Math.floor(originX / 256);
  const maxTileX = Math.floor((originX + width) / 256);
  const minTileY = Math.floor(originY / 256);
  const maxTileY = Math.floor((originY + height) / 256);

  const tilePromises = [];
  const subdomains = ['a', 'b', 'c'];
  for (let tx = minTileX; tx <= maxTileX; tx++) {
    for (let ty = minTileY; ty <= maxTileY; ty++) {
      const sub = subdomains[Math.abs(tx + ty) % 3];
      const primaryUrl = `https://${sub}.tile.openstreetmap.org/${zoom}/${tx}/${ty}.png`;
      const fallbackUrl = `https://a.tile.openstreetmap.fr/hot/${zoom}/${tx}/${ty}.png`;

      tilePromises.push(
        fetch(primaryUrl, { headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) OutdoorProposalApp/1.0' } })
          .then(async (res) => {
            if (res.ok) return res.arrayBuffer();
            const res2 = await fetch(fallbackUrl, { headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) OutdoorProposalApp/1.0' } });
            return res2.ok ? res2.arrayBuffer() : null;
          })
          .then((buf) =>
            buf
              ? new Promise((resolve) => {
                  new PNG().parse(Buffer.from(buf), (err, parsed) => {
                    resolve(err ? null : { tx, ty, img: parsed });
                  });
                })
              : null
          )
          .catch(() => null)
      );
    }
  }

  const tileResults = await Promise.all(tilePromises);
  for (const t of tileResults) {
    if (!t || !t.img) continue;
    const destX = Math.round(t.tx * 256 - originX);
    const destY = Math.round(t.ty * 256 - originY);

    for (let py = 0; py < 256; py++) {
      const cy = destY + py;
      if (cy < 0 || cy >= height) continue;
      for (let px = 0; px < 256; px++) {
        const cx = destX + px;
        if (cx < 0 || cx >= width) continue;

        const srcIdx = (256 * py + px) << 2;
        const destIdx = (width * cy + cx) << 2;

        const a = t.img.data[srcIdx + 3];
        if (a === 255) {
          canvas.data[destIdx] = t.img.data[srcIdx];
          canvas.data[destIdx + 1] = t.img.data[srcIdx + 1];
          canvas.data[destIdx + 2] = t.img.data[srcIdx + 2];
          canvas.data[destIdx + 3] = 255;
        }
      }
    }
  }

  function drawPixel(x, y, r, g, b, a = 255) {
    x = Math.round(x);
    y = Math.round(y);
    if (x < 0 || x >= width || y < 0 || y >= height) return;
    const idx = (width * y + x) << 2;
    canvas.data[idx] = r;
    canvas.data[idx + 1] = g;
    canvas.data[idx + 2] = b;
    canvas.data[idx + 3] = a;
  }

  function drawThickLine(x0, y0, x1, y1, r, g, b, thickness = 5) {
    const dx = Math.abs(x1 - x0);
    const dy = Math.abs(y1 - y0);
    const steps = Math.max(Math.ceil(Math.hypot(dx, dy)), 1);
    const half = (thickness - 1) / 2;
    for (let i = 0; i <= steps; i++) {
      const t = i / steps;
      const cx = x0 + t * (x1 - x0);
      const cy = y0 + t * (y1 - y0);
      for (let rx = -half; rx <= half; rx++) {
        for (let ry = -half; ry <= half; ry++) {
          if (rx * rx + ry * ry <= (thickness / 2) * (thickness / 2) + 0.5) {
            drawPixel(cx + rx, cy + ry, r, g, b);
          }
        }
      }
    }
  }

  const pxPoints = routePoints.map((p) => ({
    x: Math.round(lonToX(p.lng, zoom) - originX),
    y: Math.round(latToY(p.lat, zoom) - originY),
  }));

  for (let i = 0; i < pxPoints.length - 1; i++) {
    drawThickLine(pxPoints[i].x, pxPoints[i].y, pxPoints[i + 1].x, pxPoints[i + 1].y, 194, 34, 30, 5);
  }

  const clientPx = { x: Math.round(lonToX(fromLng, zoom) - originX), y: Math.round(latToY(fromLat, zoom) - originY) };
  const sitePx = { x: Math.round(lonToX(toLng, zoom) - originX), y: Math.round(latToY(toLat, zoom) - originY) };

  const { greenPin, redPin } = getPinIcons();

  // Draw ONLY pin_green at client location and pin_red_adinn at site location (no extra or last pins)
  overlayPinImage(canvas, greenPin, clientPx.x, clientPx.y, 0.22);
  overlayPinImage(canvas, redPin, sitePx.x, sitePx.y, 0.22);

  const buffer = PNG.sync.write(canvas);
  return { buffer, distanceText, durationText };
}

/**
 * Fetches or generates a route map image (PNG buffer) between two points, plus the route's
 * driving distance/duration text (e.g. "3.5 km", "4 mins").
 * Renders pin_green at client location and pin_red_adinn at site location with zoomed-out framing
 * so the full route and all pin markers are clearly visible inside the map image.
 */
async function getRouteMapBuffer({ fromLat, fromLng, toLat, toLng, width = 640, height = 480 }) {
  const fLat = Number(fromLat);
  const fLng = Number(fromLng);
  const tLat = Number(toLat);
  const tLng = Number(toLng);

  if (
    !Number.isFinite(fLat) ||
    !Number.isFinite(fLng) ||
    !Number.isFinite(tLat) ||
    !Number.isFinite(tLng) ||
    fLat < -90 ||
    fLat > 90 ||
    tLat < -90 ||
    tLat > 90 ||
    fLng < -180 ||
    fLng > 180 ||
    tLng < -180 ||
    tLng > 180 ||
    (fLat === 0 && fLng === 0) ||
    (tLat === 0 && tLng === 0)
  ) {
    return null;
  }

  return renderRouteMapImage({ fromLat: fLat, fromLng: fLng, toLat: tLat, toLng: tLng, width, height });
}

module.exports = { getRouteMapBuffer };
