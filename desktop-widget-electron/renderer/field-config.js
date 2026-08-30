// Per-tab field behavior: which fields get dropdowns, a date picker, or an
// auto-computed "Week" — everything else falls back to a plain text box.
// Tab names here must match the actual sheet tab names exactly.

const SITES = ['MHF', 'FV', 'Peds', 'GI', 'Specialty Pharmacy', 'All'];
const REQUESTERS = ['Cassandra', 'Tseten', 'Grant', 'Tammy', 'Naveen'];
const ASSIGNEES = ['Naveen', 'Surya', 'Amulya', 'Cassandra', 'Tseten', 'Grant', 'Tammy', 'Lucy', 'Erika'];

// Top-level menu the widget shows before the tab list — started life as
// just "log a status update", now covers more than one kind of thing, so
// tabs group under whichever menu they conceptually belong to instead of
// one long flat list. A tab that shows up from the sheet but isn't listed
// under any of these still appears — see groupTabsIntoCategories() below —
// bucketed into an auto-created "More" menu, so a newly added tab is
// never silently unreachable, just uncategorized until it's added here.
const CATEGORIES = {
  'Status Report Generator': ['Daily Status', 'Adhoc_Mails', 'Cleanup_Activities', 'Drupal_Bugs_&_Improvements'],
  'Team Management': ['Leave'],
};

const FIELD_CONFIG = {
  'Daily Status': {
    Site: { type: 'select', options: SITES },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    Status: { type: 'select', options: ['Done', 'Pending', 'In Progress', 'Open'] },
    Priority: { type: 'select', options: ['High', 'Medium', 'Low'] },
    'Assigned to': { type: 'multiselect', options: ASSIGNEES },
  },
  Adhoc_Mails: {
    Requester: { type: 'select', options: REQUESTERS },
    Site: { type: 'select', options: SITES },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    'Assigned to': { type: 'multiselect', options: ASSIGNEES },
  },
  Cleanup_Activities: {
    'Cleanup Number': { type: 'sequence' },
    Volume: { type: 'select', options: ['Large', 'Medium', 'Small'] },
    Requester: { type: 'select', options: REQUESTERS },
    'Site Impacted': { type: 'select', options: SITES },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    'Assigned to': { type: 'multiselect', options: ASSIGNEES },
  },
  'Drupal_Bugs_&_Improvements': {
    Website: { type: 'select', options: SITES },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    Type: { type: 'select', options: ['Bug', 'Fix', 'Suggestion'] },
    'Assigned to': { type: 'multiselect', options: ASSIGNEES },
  },
  // Auto-created server-side (see ensureLeaveTab() in Code.gs) the first
  // time the widget asks for the tab list — this config just makes that
  // tab's form as polished as the rest instead of falling back to plain
  // text boxes for every field. One person per leave entry, so a plain
  // `select` (not `multiselect` like "Assigned to" elsewhere) fits better.
  Leave: {
    Name: { type: 'select', options: ASSIGNEES },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    Type: { type: 'select', options: ['Vacation', 'Sick', 'WFH', 'Holiday', 'Other'] },
  },
};

function normalizeKey(s) {
  return String(s).trim().toLowerCase();
}

// Buckets the live list of sheet tabs into CATEGORIES, in CATEGORIES'
// own order, skipping a category entirely if none of its tabs currently
// exist (e.g. Leave got renamed). Anything left over — present on the
// sheet but not listed under any category — lands in a "More" bucket
// appended at the end, shown only if it's non-empty, so a brand-new tab
// is discoverable instead of just vanishing from the widget.
function groupTabsIntoCategories(tabs) {
  const used = new Set();
  const groups = Object.keys(CATEGORIES)
    .map((name) => {
      const members = CATEGORIES[name].filter((t) => tabs.indexOf(t) !== -1);
      members.forEach((t) => used.add(t));
      return { name: name, tabs: members };
    })
    .filter((g) => g.tabs.length > 0);
  const leftover = tabs.filter((t) => !used.has(t));
  if (leftover.length) groups.push({ name: 'More', tabs: leftover });
  return groups;
}

// Matches column names case-insensitively (and ignoring stray whitespace),
// since the sheet's actual header casing ("Assigned To" vs "Assigned to")
// won't always match this config exactly.
function fieldSpecFor(tab, column) {
  const tabConfig = FIELD_CONFIG[tab];
  if (!tabConfig) return null;
  const target = normalizeKey(column);
  const matchKey = Object.keys(tabConfig).find((k) => normalizeKey(k) === target);
  return matchKey ? tabConfig[matchKey] : null;
}

// "Week N" of the MONTH, counting only Mon-Fri work weeks (weekends are
// off days and don't start a new week on their own). Week 1 starts on the
// 1st if it's a weekday, otherwise on the first Monday on/after the 1st —
// e.g. Aug 2026: Aug 1 is a Saturday, so Week 1 = Aug 3-7 (Mon-Fri),
// Week 2 = Aug 10-14, and so on.
function computeWeekLabel(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  if (isNaN(d.getTime())) return '';

  const year = d.getFullYear();
  const month = d.getMonth();
  const first = new Date(year, month, 1);
  const firstDow = first.getDay(); // 0=Sun..6=Sat

  let week1Start;
  if (firstDow === 0) week1Start = new Date(year, month, 2); // Sun -> Mon
  else if (firstDow === 6) week1Start = new Date(year, month, 3); // Sat -> Mon
  else week1Start = first; // month already starts on a weekday

  if (d < week1Start) return 'Week 1'; // a leave day before the month's first work week

  const week1Dow = (week1Start.getDay() + 6) % 7; // Monday = 0
  const daysToNextMonday = 7 - week1Dow;
  const week2Start = new Date(week1Start);
  week2Start.setDate(week1Start.getDate() + daysToNextMonday);

  if (d < week2Start) return 'Week 1';

  const daysSinceWeek2 = Math.round((d - week2Start) / 86400000);
  return `Week ${2 + Math.floor(daysSinceWeek2 / 7)}`;
}
