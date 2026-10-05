// Illumination as shown in Excel files — "Not Lit" (any spacing/case) is shown as "Non Lit".
function illuminationLabel(value) {
  if (!value) return value;
  return /^\s*not\s*-?\s*lit\s*$/i.test(value) ? 'Non Lit' : value;
}

module.exports = { illuminationLabel };
