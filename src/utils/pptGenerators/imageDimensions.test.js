import { describe, it, expect } from 'vitest';
import { getImageDimensions } from './imageDimensions.js';

// 1x1 PNG
const PNG_1x1 = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64');
// 2x1 baseline JPEG
const JPEG_2x1 = Buffer.from(
  '/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/wAALCAABAAIBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAAB//EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAT8AVN//2Q==',
  'base64'
);

describe('getImageDimensions', () => {
  it('reads PNG width/height from the IHDR chunk', () => {
    expect(getImageDimensions(PNG_1x1, 'png')).toEqual({ width: 1, height: 1 });
  });

  it('reads JPEG width/height from the SOF marker', () => {
    const dims = getImageDimensions(JPEG_2x1, 'jpg');
    expect(dims).toBeTruthy();
    expect(dims.width).toBeGreaterThan(0);
    expect(dims.height).toBeGreaterThan(0);
  });

  it('returns null for unsupported/empty/malformed input instead of throwing', () => {
    expect(getImageDimensions(null, 'png')).toBeNull();
    expect(getImageDimensions(Buffer.from('not an image'), 'png')).toBeNull();
    expect(getImageDimensions(Buffer.alloc(0), 'jpg')).toBeNull();
    expect(getImageDimensions(Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"></svg>'), 'svg')).toBeNull();
  });
});
