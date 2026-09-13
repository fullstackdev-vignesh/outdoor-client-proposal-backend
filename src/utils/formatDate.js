const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// Site/Booking dates are pre-shifted by +5:30 (IST) before being stored (see nowIST() in
// models/controllers), so reading them back must use the UTC getters to avoid re-applying
// the server's local timezone offset on top of the already-shifted value.
function formatIST(value) {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return '';
  const day = String(d.getUTCDate()).padStart(2, '0');
  const month = MONTHS[d.getUTCMonth()];
  const year = d.getUTCFullYear();
  let hours = d.getUTCHours();
  const minutes = String(d.getUTCMinutes()).padStart(2, '0');
  const ampm = hours >= 12 ? 'pm' : 'am';
  hours = hours % 12 || 12;
  return `${day}/${month}/${year}, ${hours}:${minutes} ${ampm}`;
}

module.exports = { formatIST };
