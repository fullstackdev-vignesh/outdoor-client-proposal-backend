import { describe, it, expect } from 'vitest';
import { resolveGenerator } from './index.js';

describe('resolveGenerator - template name -> generator mapping', () => {
  it('resolves the Adinn New Template generator (case-insensitive, trimmed)', () => {
    expect(resolveGenerator('Adinn New Template')).toBeTypeOf('function');
    expect(resolveGenerator('  adinn new template  ')).toBeTypeOf('function');
    expect(resolveGenerator('ADINN NEW TEMPLATE')).toBeTypeOf('function');
  });

  it('returns null for other/unknown PPT templates so they keep using the legacy engine', () => {
    expect(resolveGenerator('PPT Master Template')).toBeNull();
    expect(resolveGenerator('Some Other Template')).toBeNull();
    expect(resolveGenerator(undefined)).toBeNull();
    expect(resolveGenerator(null)).toBeNull();
    expect(resolveGenerator('')).toBeNull();
  });
});
