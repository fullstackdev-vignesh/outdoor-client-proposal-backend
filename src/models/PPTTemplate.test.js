import { describe, it, expect } from 'vitest';
import PPTTemplate from './PPTTemplate.js';

describe('PPTTemplate model - PPT Master create requirements', () => {
  it('validates with only name, description and status (no version/variant/file)', () => {
    const doc = new PPTTemplate({ name: 'Q1 Proposal', description: 'Q1 pitch deck', status: 'active' });
    const err = doc.validateSync();
    expect(err).toBeUndefined();
  });

  it('rejects a missing Template Name', () => {
    const doc = new PPTTemplate({ description: 'no name', status: 'active' });
    const err = doc.validateSync();
    expect(err).toBeDefined();
    expect(err.errors.name).toBeDefined();
  });

  it('defaults version and variant when not supplied (not required for create)', () => {
    const doc = new PPTTemplate({ name: 'Q1 Proposal' });
    const err = doc.validateSync();
    expect(err).toBeUndefined();
    expect(doc.version).toBe('1.0');
    expect(doc.variant).toBe('Standard');
  });

  it('does not require a file (fileUrl optional)', () => {
    const doc = new PPTTemplate({ name: 'Q1 Proposal' });
    const err = doc.validateSync();
    expect(err).toBeUndefined();
    expect(doc.fileUrl).toBeUndefined();
  });

  it('still accepts explicit version/variant/fileUrl for existing records (edit flow)', () => {
    const doc = new PPTTemplate({ name: 'Legacy', version: '2.3', variant: 'Premium', fileUrl: '/uploads/old.pptx' });
    const err = doc.validateSync();
    expect(err).toBeUndefined();
    expect(doc.version).toBe('2.3');
    expect(doc.variant).toBe('Premium');
    expect(doc.fileUrl).toBe('/uploads/old.pptx');
  });
});
