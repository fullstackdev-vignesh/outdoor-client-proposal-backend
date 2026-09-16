const EARTH_RADIUS_KM = 6371;

function toRad(deg) {
  return (deg * Math.PI) / 180;
}

/**
 * Straight-line (great-circle) distance in KM between two lat/lng points.
 * Returns null for missing/invalid coordinates instead of throwing.
 */
function haversineDistanceKm(lat1, lng1, lat2, lng2) {
  const raw = [lat1, lng1, lat2, lng2];
  if (raw.some((v) => v === undefined || v === null || v === '' || Number.isNaN(Number(v)))) {
    return null;
  }
  const [a1, o1, a2, o2] = raw.map(Number);
  if ([a1, a2].some((v) => v < -90 || v > 90) || [o1, o2].some((v) => v < -180 || v > 180)) {
    return null;
  }

  const dLat = toRad(a2 - a1);
  const dLng = toRad(o2 - o1);
  const sinDLat = Math.sin(dLat / 2);
  const sinDLng = Math.sin(dLng / 2);
  const h = sinDLat * sinDLat + Math.cos(toRad(a1)) * Math.cos(toRad(a2)) * sinDLng * sinDLng;
  const c = 2 * Math.atan2(Math.sqrt(h), Math.sqrt(1 - h));
  return EARTH_RADIUS_KM * c;
}

module.exports = { haversineDistanceKm };
