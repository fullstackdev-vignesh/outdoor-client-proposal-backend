const adinnNewTemplate = require('./adinnNewTemplate');

// Template name (case-insensitive, trimmed) -> generator function.
// Add a new entry here for each future PPT template that needs its own
// independent slide design; templates not listed fall back to the legacy
// PptxTemplate (master.pptx) engine, so existing templates are unaffected.
const GENERATORS = {
  'adinn new template': adinnNewTemplate.generate,
};

function resolveGenerator(templateName) {
  if (!templateName || typeof templateName !== 'string') return null;
  return GENERATORS[templateName.trim().toLowerCase()] || null;
}

module.exports = { resolveGenerator, GENERATORS };
