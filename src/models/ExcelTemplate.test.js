import { describe, it, expect } from 'vitest';
import ExcelTemplate from './ExcelTemplate.js';

describe('ExcelTemplate model - Excel Template create requirements', () => {
  it('validates with only name, description and status (no version/file)', () => {
    const doc = new ExcelTemplate({ name: 'Q1 Excel Export', description: 'Q1 rate card', status: 'active' });
    const err = doc.validateSync();
    expect(err).toBeUndefined();
  });

  it('rejects a missing Template Name', () => {
    const doc = new ExcelTemplate({ description: 'no name', status: 'active' });
    const err = doc.validateSync();
    expect(err).toBeDefined();
    expect(err.errors.name).toBeDefined();
  });

  it('defaults version when not supplied (not required for create)', () => {
    const doc = new ExcelTemplate({ name: 'Q1 Excel Export' });
    const err = doc.validateSync();
    expect(err).toBeUndefined();
    expect(doc.version).toBe('1.0');
  });

  it('does not require a file (fileUrl optional)', () => {
    const doc = new ExcelTemplate({ name: 'Q1 Excel Export' });
    const err = doc.validateSync();
    expect(err).toBeUndefined();
    expect(doc.fileUrl).toBeUndefined();
  });

  it('still accepts explicit version/fileUrl for existing records (edit flow)', () => {
    const doc = new ExcelTemplate({ name: 'Legacy', version: '2.3', fileUrl: '/uploads/old.xlsx' });
    const err = doc.validateSync();
    expect(err).toBeUndefined();
    expect(doc.version).toBe('2.3');
    expect(doc.fileUrl).toBe('/uploads/old.xlsx');
  });
});
