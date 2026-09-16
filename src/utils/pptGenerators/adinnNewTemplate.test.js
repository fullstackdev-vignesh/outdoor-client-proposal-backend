import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { generate } from './adinnNewTemplate.js';

const TEMPLATE_PATH = path.join(__dirname, '..', '..', '..', 'assets', 'proposal-templates', 'adinn-new-template.pptx');

// 1x1 transparent PNG, used as a deterministic stand-in for a real site/map image.
const TINY_PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
const TINY_PNG_BUFFER = Buffer.from(TINY_PNG_BASE64, 'base64');

// generate() takes an optional `deps` override for its map/image fetchers instead
// of relying on vi.mock — this file (and pptxTemplateEngine.js) is plain CJS
// require() with no ESM syntax, and Vitest's mock interception does not reach
// into a CJS module's own internal require() calls, so vi.mock silently no-ops
// here. Dependency injection is the reliable way to stub network/disk I/O.
const noNetworkDeps = { getRouteMapBuffer: vi.fn(async () => null), getImageBuffer: vi.fn(async () => null) };

function gen({ deps, ...args }) {
  return generate({ deps: { ...noNetworkDeps, ...deps }, ...args });
}

async function loadZip(buffer) {
  return JSZip.loadAsync(buffer);
}

/** Resolves the slide XMLs in true presentation order (via presentation.xml's
 * sldIdLst + presentation.xml.rels), not by slideN.xml filename — cloned site
 * slides are named slide_gen_*.xml and do not sort correctly by number. */
async function getOrderedSlideXmls(buffer) {
  const zip = await loadZip(buffer);
  const pres = await zip.file('ppt/presentation.xml').async('string');
  const presRels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  const relIdToTarget = {};
  for (const m of presRels.matchAll(/<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relIdToTarget[m[1]] = m[2];
  }
  const sldIdLst = pres.match(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/)[0];
  const relIds = [...sldIdLst.matchAll(/r:id="([^"]+)"/g)].map((m) => m[1]);
  return Promise.all(relIds.map((relId) => zip.file(`ppt/${relIdToTarget[relId]}`).async('string')));
}

async function getMediaFileNames(buffer) {
  const zip = await loadZip(buffer);
  return Object.keys(zip.files).filter((f) => /^ppt\/media\/.+\.(png|jpe?g|gif|webp|svg)$/i.test(f));
}

function textsOf(xml) {
  return [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]);
}

function site(overrides = {}) {
  return {
    mediaId: 'MED1',
    city: 'Salem',
    state: 'TN',
    location: 'Anna Statue Junction',
    width: 30,
    height: 15,
    mediaType: 'Unipole',
    illumination: 'Backlit',
    quantity: 3,
    latitude: 11.65,
    longitude: 78.15,
    mediaImage: null,
    ...overrides,
  };
}

const client = { name: 'Test Client XYZ', customerType: 'client', latitude: 11.0168, longitude: 76.9558 };

beforeEach(() => {
  // Only fake Date (not timers) — pptxTemplateEngine/JSZip rely on real
  // setImmediate/setTimeout scheduling internally, and faking those hangs save().
  vi.useFakeTimers({ toFake: ['Date'] });
  vi.setSystemTime(new Date(2026, 7, 10)); // Aug 10, 2026
});

afterEach(() => {
  vi.useRealTimers();
});

describe('Adinn New Template PPT generator - source template integrity', () => {
  it('never modifies the original template file on disk', async () => {
    const before = crypto.createHash('sha256').update(fs.readFileSync(TEMPLATE_PATH)).digest('hex');
    await gen({ proposal: { client }, client, sites: [site(), site({ mediaId: 'MED2' })] });
    const after = crypto.createHash('sha256').update(fs.readFileSync(TEMPLATE_PATH)).digest('hex');
    expect(after).toBe(before);
  });

  it('generated file opens successfully (valid zip/XML round-trip)', async () => {
    const buf = await gen({ proposal: { client }, client, sites: [site()] });
    const zip = await loadZip(buf);
    expect(zip.file('ppt/presentation.xml')).toBeTruthy();
    const texts = await getOrderedSlideXmls(buf);
    expect(texts.every((xml) => xml.startsWith('<?xml'))).toBe(true);
  });
});

describe('Adinn New Template PPT generator - slide count', () => {
  it('1 site => 6 slides (4 fixed + 2)', async () => {
    const buf = await gen({ proposal: { client }, client, sites: [site()] });
    const texts = await getOrderedSlideXmls(buf);
    expect(texts).toHaveLength(6);
  });

  it('2 sites => 8 slides (4 fixed + 4)', async () => {
    const buf = await gen({ proposal: { client }, client, sites: [site({ mediaId: 'MED1' }), site({ mediaId: 'MED2' })] });
    const texts = await getOrderedSlideXmls(buf);
    expect(texts).toHaveLength(8);
  });

  it('N sites => 4 + 2*N slides', async () => {
    const sites = [site({ mediaId: 'A' }), site({ mediaId: 'B' }), site({ mediaId: 'C' })];
    const buf = await gen({ proposal: { client }, client, sites });
    const texts = await getOrderedSlideXmls(buf);
    expect(texts).toHaveLength(4 + 2 * sites.length);
  });
});

describe('Adinn New Template PPT generator - dynamic content', () => {
  it('replaces Slide 1 date and client/agency name', async () => {
    const buf = await gen({ proposal: { client }, client, sites: [site()] });
    const [coverXml] = await getOrderedSlideXmls(buf);
    expect(textsOf(coverXml)).toContain('ate: Aug 10, 2026');
    expect(textsOf(coverXml)).toContain('Test Client XYZ');
    expect(textsOf(coverXml)).not.toContain('Maxi Vision Eye Hospital');
  });

  it('leaves Slide 2 (About Us) and Slide 3 (Why Adinn) text fully intact', async () => {
    const buf = await gen({ proposal: { client }, client, sites: [site()] });
    const [, aboutXml, whyXml] = await getOrderedSlideXmls(buf);
    expect(textsOf(aboutXml)).toContain('ABOUT US');
    expect(textsOf(aboutXml)).toContain('Who we are');
    expect(textsOf(whyXml)).toContain('Why Leading Brands Choose Adinn Advertising');
    expect(textsOf(whyXml)).toContain('Prime, High-Impact Locations');
  });

  it('leaves the final Thank You slide unchanged and last regardless of site count', async () => {
    const sites = [site({ mediaId: 'A' }), site({ mediaId: 'B' })];
    const buf = await gen({ proposal: { client }, client, sites });
    const texts = await getOrderedSlideXmls(buf);
    const thankYou = texts[texts.length - 1];
    expect(textsOf(thankYou)).toContain('Thank You');
    expect(textsOf(thankYou)).toContain('Adinn Outdoor Advertising');
    expect(textsOf(thankYou)).toContain('Awaiting Your Approval...');
  });

  it('maps correct site data onto each duplicated Spec+Map slide pair, in selection order', async () => {
    const siteA = site({ mediaId: 'A', city: 'Salem', width: 30, height: 15, mediaType: 'Unipole', illumination: 'Backlit', quantity: 3, location: 'Anna Statue Junction' });
    const siteB = site({ mediaId: 'B', city: 'Coimbatore', width: 20, height: 10, mediaType: 'Hoarding', illumination: 'Frontlit', quantity: 1, location: 'Avinashi Road' });
    const buf = await gen({ proposal: { client }, client, sites: [siteA, siteB] });
    const texts = await getOrderedSlideXmls(buf);
    // 0 cover, 1 about, 2 why, 3 A-spec, 4 A-map, 5 B-spec, 6 B-map, 7 thankyou
    const aSpec = textsOf(texts[3]);
    expect(aSpec).toContain('Salem');
    expect(aSpec).toContain('30x15');
    expect(aSpec).toContain('Unipole');
    expect(aSpec).toContain('Backlit');
    expect(aSpec).toContain('3');
    expect(aSpec).toContain('Anna Statue Junction 30x15');
    expect(textsOf(texts[4])).toContain('Anna Statue Junction 30x15');

    const bSpec = textsOf(texts[5]);
    expect(bSpec).toContain('Coimbatore');
    expect(bSpec).toContain('20x10');
    expect(bSpec).toContain('Hoarding');
    expect(bSpec).toContain('Frontlit');
    expect(bSpec).toContain('Avinashi Road 20x10');
    expect(textsOf(texts[6])).toContain('Avinashi Road 20x10');
  });
});

describe('Adinn New Template PPT generator - site photo replacement', () => {
  it('inserts the site photo as an editable native picture on the spec slide, sized to the existing photo area', async () => {
    const getImageBuffer = vi.fn(async () => ({ buffer: TINY_PNG_BUFFER, ext: 'png' }));
    const buf = await gen({ proposal: { client }, client, sites: [site({ mediaImage: '/uploads/site1.png' })], deps: { getImageBuffer } });
    const texts = await getOrderedSlideXmls(buf);
    const specXml = texts[3];
    expect(specXml).toMatch(/<p:pic>/);
    expect(getImageBuffer).toHaveBeenCalledWith('/uploads/site1.png');
    const media = await getMediaFileNames(buf);
    expect(media.length).toBeGreaterThan(0);
  });

  it('reuses the same site photo on the map slide photo slot (existing shape, not a raster full-slide swap)', async () => {
    const getImageBuffer = vi.fn(async () => ({ buffer: TINY_PNG_BUFFER, ext: 'png' }));
    const buf = await gen({ proposal: { client }, client, sites: [site({ mediaImage: '/uploads/site1.png' })], deps: { getImageBuffer } });
    const texts = await getOrderedSlideXmls(buf);
    const mapXml = texts[4];
    // the map slide's photo slot is an existing template shape (blipFill), not a new <p:pic>
    expect(mapXml).toContain('r:embed="rId6"');
    expect(mapXml).not.toMatch(/<p:pic>/);
  });

  it('falls back to a placeholder (no crash) when no site photo is available', async () => {
    const buf = await gen({ proposal: { client }, client, sites: [site({ mediaImage: null })] });
    const texts = await getOrderedSlideXmls(buf);
    expect(textsOf(texts[3])).toContain('Site image not available');
    expect(texts[3]).not.toMatch(/<p:pic>/);
  });
});

describe('Adinn New Template PPT generator - map/distance', () => {
  it('computes and renders a valid Haversine distance when coordinates are valid', async () => {
    const buf = await gen({ proposal: { client }, client, sites: [site({ latitude: 11.65, longitude: 78.15 })] });
    const texts = await getOrderedSlideXmls(buf);
    expect(textsOf(texts[4])).toEqual(expect.arrayContaining([expect.stringMatching(/Distance: \d+\.\d km/)]));
  });

  it('embeds the generated route-map image into the existing map slot when coordinates resolve', async () => {
    const getRouteMapBuffer = vi.fn(async () => TINY_PNG_BUFFER);
    const buf = await gen({ proposal: { client }, client, sites: [site({ latitude: 11.65, longitude: 78.15 })], deps: { getRouteMapBuffer } });
    const media = await getMediaFileNames(buf);
    expect(media.length).toBeGreaterThan(0);
    expect(getRouteMapBuffer).toHaveBeenCalledWith(
      expect.objectContaining({ fromLat: client.latitude, fromLng: client.longitude, toLat: 11.65, toLng: 78.15 })
    );
  });

  it('handles missing/invalid site or client coordinates safely (no crash, distance N/A)', async () => {
    const buf = await gen({ proposal: { client }, client, sites: [site({ latitude: null, longitude: null })] });
    const texts = await getOrderedSlideXmls(buf);
    expect(textsOf(texts[4])).toContain('Distance: N/A');
  });
});
