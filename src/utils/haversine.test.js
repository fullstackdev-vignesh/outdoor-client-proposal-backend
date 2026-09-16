import { describe, it, expect } from 'vitest';
import { haversineDistanceKm } from './haversine.js';

describe('haversineDistanceKm', () => {
  it('returns a valid KM distance for two known coordinates', () => {
    // Chennai (13.0827, 80.2707) to Bengaluru (12.9716, 77.5946) ~ 290km straight-line
    const km = haversineDistanceKm(13.0827, 80.2707, 12.9716, 77.5946);
    expect(km).not.toBeNull();
    expect(km).toBeGreaterThan(280);
    expect(km).toBeLessThan(300);
  });

  it('returns 0 for identical coordinates', () => {
    expect(haversineDistanceKm(11.0168, 76.9558, 11.0168, 76.9558)).toBeCloseTo(0, 5);
  });

  it('returns null when coordinates are missing', () => {
    expect(haversineDistanceKm(undefined, undefined, 11.0168, 76.9558)).toBeNull();
    expect(haversineDistanceKm(null, 76.9558, 11.0168, 76.9558)).toBeNull();
  });

  it('returns null for invalid/out-of-range coordinates', () => {
    expect(haversineDistanceKm('abc', 76.9558, 11.0168, 76.9558)).toBeNull();
    expect(haversineDistanceKm(999, 76.9558, 11.0168, 76.9558)).toBeNull();
    expect(haversineDistanceKm(11.0168, 999, 11.0168, 76.9558)).toBeNull();
  });
});
