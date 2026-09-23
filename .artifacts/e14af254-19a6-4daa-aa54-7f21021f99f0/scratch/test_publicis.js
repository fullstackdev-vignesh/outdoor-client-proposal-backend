function lonToX(lon, zoom) {
  return ((lon + 180) / 360) * Math.pow(2, zoom) * 256;
}

function latToY(lat, zoom) {
  const sin = Math.sin((lat * Math.PI) / 180);
  const rad = Math.log((1 + sin) / (1 - sin)) / 2;
  return (0.5 - rad / (2 * Math.PI)) * Math.pow(2, zoom) * 256;
}

// Thirumazhisai -> SHU88 / Marina Beach (from user's screenshot: 26.3 km, 24 mins)
const fromLat = 13.030, fromLng = 80.060; // Thirumazhisai
const toLat = 13.050, toLng = 80.280;   // SHU88 / Marina Beach

const width = 640, height = 480;

const minLat = Math.min(fromLat, toLat);
const maxLat = Math.max(fromLat, toLat);
const minLng = Math.min(fromLng, toLng);
const maxLng = Math.max(fromLng, toLng);

console.log('Testing zoom levels for 26.3 km route:');
for (let z = 15; z >= 9; z--) {
  const x1 = lonToX(minLng, z);
  const y1 = latToY(maxLat, z);
  const x2 = lonToX(maxLng, z);
  const y2 = latToY(minLat, z);
  const dx = Math.abs(x2 - x1);
  const dy = Math.abs(y2 - y1);

  const centerX = (lonToX(minLng, z) + lonToX(maxLng, z)) / 2;
  const centerY = (latToY(minLat, z) + latToY(maxLat, z)) / 2;
  const originX = centerX - width / 2;
  const originY = centerY - height / 2;

  const p1X = lonToX(fromLng, z) - originX;
  const p1Y = latToY(fromLat, z) - originY;
  const p2X = lonToX(toLng, z) - originX;
  const p2Y = latToY(toLat, z) - originY;

  console.log(`Zoom ${z}: dx=${dx.toFixed(1)}, dy=${dy.toFixed(1)} | Pin1 X=${p1X.toFixed(1)}, Pin2 X=${p2X.toFixed(1)}`);
}
