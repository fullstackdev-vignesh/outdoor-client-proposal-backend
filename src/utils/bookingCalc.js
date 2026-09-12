const MS_PER_DAY = 24 * 60 * 60 * 1000;

function calcDurationDays(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const diff = Math.round((end - start) / MS_PER_DAY);
  return diff >= 0 ? diff + 1 : 0;
}

function calcBookingAmount(monthlyTotalCost, durationDays) {
  const cost = Number(monthlyTotalCost) || 0;
  const days = Number(durationDays) || 0;
  return Math.round(((cost / 30) * days + Number.EPSILON) * 100) / 100;
}

module.exports = { calcDurationDays, calcBookingAmount };
