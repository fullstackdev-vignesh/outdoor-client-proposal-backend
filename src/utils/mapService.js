const GOOGLE_STATIC_MAP_URL = 'https://maps.googleapis.com/maps/api/staticmap';
const GOOGLE_DIRECTIONS_URL = 'https://maps.googleapis.com/maps/api/directions/json';

/**
 * Fetches a real Google route map image (PNG buffer) between two points, plus the route's
 * driving distance/duration text (e.g. "6.8 km", "19 mins") — via the Google Directions API
 * (route + distance/duration) and the Google Static Maps API (the rendered image, path +
 * markers). Requires GOOGLE_MAPS_API_KEY to be set in the environment; returns null if the key
 * is missing, coordinates are missing/invalid, or either Google API call fails (e.g. no route,
 * no network, billing not enabled on the key) — callers fall back to the usual "Insert your map
 * image here" placeholder in that case.
 */
async function getRouteMapBuffer({ fromLat, fromLng, toLat, toLng, width = 640, height = 480 }) {
  const apiKey = process.env.GOOGLE_MAPS_API_KEY;
  if (!apiKey) return null;
  if ([fromLat, fromLng, toLat, toLng].some((v) => v === undefined || v === null || Number.isNaN(Number(v)))) {
    return null;
  }
  try {
    const dirParams = new URLSearchParams({
      origin: `${fromLat},${fromLng}`,
      destination: `${toLat},${toLng}`,
      mode: 'driving',
      key: apiKey,
    });
    const dirRes = await fetch(`${GOOGLE_DIRECTIONS_URL}?${dirParams.toString()}`);
    if (!dirRes.ok) return null;
    const dirData = await dirRes.json();
    const route = dirData?.routes?.[0];
    const leg = route?.legs?.[0];
    if (!route || !leg) return null;

    const distanceText = leg.distance?.text || null;
    const durationText = leg.duration?.text || null;
    const overviewPolyline = route.overview_polyline?.points;

    const mapParams = new URLSearchParams({ size: `${width}x${height}`, key: apiKey });
    mapParams.append('markers', `color:green|label:A|${fromLat},${fromLng}`);
    mapParams.append('markers', `color:red|label:B|${toLat},${toLng}`);
    if (overviewPolyline) {
      mapParams.append('path', `color:0xC2221EDD|weight:4|enc:${overviewPolyline}`);
    }

    const mapRes = await fetch(`${GOOGLE_STATIC_MAP_URL}?${mapParams.toString()}`);
    if (!mapRes.ok) return null;
    const arrayBuffer = await mapRes.arrayBuffer();
    return { buffer: Buffer.from(arrayBuffer), distanceText, durationText };
  } catch {
    return null;
  }
}

module.exports = { getRouteMapBuffer };
