const OSRM_URL = 'https://router.project-osrm.org/route/v1/driving';
const STATICMAP_URL = 'https://staticmap.openstreetmap.de/staticmap.php';

function decodePolyline(str) {
  let index = 0;
  let lat = 0;
  let lng = 0;
  const points = [];
  while (index < str.length) {
    let result = 1;
    let shift = 0;
    let b;
    do {
      b = str.charCodeAt(index++) - 63 - 1;
      result += b << shift;
      shift += 5;
    } while (b >= 0x1f);
    lat += result & 1 ? ~(result >> 1) : result >> 1;

    result = 1;
    shift = 0;
    do {
      b = str.charCodeAt(index++) - 63 - 1;
      result += b << shift;
      shift += 5;
    } while (b >= 0x1f);
    lng += result & 1 ? ~(result >> 1) : result >> 1;

    points.push([lat * 1e-5, lng * 1e-5]);
  }
  return points;
}

function samplePoints(points, maxPoints = 80) {
  if (points.length <= maxPoints) return points;
  const step = points.length / maxPoints;
  const sampled = [];
  for (let i = 0; i < maxPoints; i++) sampled.push(points[Math.floor(i * step)]);
  sampled.push(points[points.length - 1]);
  return sampled;
}

/**
 * Fetches a static route map image (PNG buffer) between two points using free,
 * keyless services: OSRM demo server for routing and staticmap.openstreetmap.de
 * for rendering. Returns null if either lookup fails (e.g. no network, no route).
 */
async function getRouteMapBuffer({ fromLat, fromLng, toLat, toLng, width = 640, height = 480 }) {
  if ([fromLat, fromLng, toLat, toLng].some((v) => v === undefined || v === null || Number.isNaN(Number(v)))) {
    return null;
  }
  try {
    const routeRes = await fetch(`${OSRM_URL}/${fromLng},${fromLat};${toLng},${toLat}?overview=full&geometries=polyline`);
    if (!routeRes.ok) return null;
    const routeData = await routeRes.json();
    const geometry = routeData?.routes?.[0]?.geometry;
    if (!geometry) return null;

    const points = samplePoints(decodePolyline(geometry));
    const pathParam = `color:0xC2221EDD|weight:4|${points.map(([lat, lng]) => `${lat.toFixed(5)},${lng.toFixed(5)}`).join('|')}`;

    const params = new URLSearchParams();
    params.set('size', `${width}x${height}`);
    params.set('maptype', 'mapnik');
    params.append('path', pathParam);
    params.append('markers', `color:green|label:A|${fromLat},${fromLng}`);
    params.append('markers', `color:red|label:B|${toLat},${toLng}`);

    const mapRes = await fetch(`${STATICMAP_URL}?${params.toString()}`);
    if (!mapRes.ok) return null;
    const arrayBuffer = await mapRes.arrayBuffer();
    return Buffer.from(arrayBuffer);
  } catch {
    return null;
  }
}

module.exports = { getRouteMapBuffer, decodePolyline };
