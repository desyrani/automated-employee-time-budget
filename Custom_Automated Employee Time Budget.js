// return first monday relative to first day of the month until sunday of current week
function getDefaultDate() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const date = now.getDate();
  const dayOfWeek = now.getDay();
  const offSet = dayOfWeek === 0 ? dayOfWeek + 7 : dayOfWeek;
  const thisMonth = new Date(year, month, 1);
  const firstDate = new Date(year, month, 1 - (thisMonth.getDay() - 1));
  const lastDate = new Date(year, month, date + 7 - offSet, 23, 59);

  return { firstDate, lastDate };
}

// return ISOWeekNumber of a given date
function getISOWeekNumber(date) {
  let target = new Date(date.valueOf());
  const dayNumber = (date.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNumber + 3);
  const firstThursday = target.valueOf();
  target.setMonth(0, 1);

  if (target.getDay() !== 4) {
    target.setMonth(0, 1 + ((4 - target.getDay() + 7) % 7));
  }

  return 1 + Math.ceil((firstThursday - target) / (7 * 24 * 3600 * 1000));
}
