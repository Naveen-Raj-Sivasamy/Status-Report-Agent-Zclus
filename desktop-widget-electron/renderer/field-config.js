// Per-tab field behavior: which fields get dropdowns, a date picker, or an
// auto-computed "Week" — everything else falls back to a plain text box.
// Tab names here must match the actual sheet tab names exactly.

// Every dropdown/multiselect below references an *options key* (a string,
// e.g. 'SITES' or 'Daily Status.Status') instead of a literal array — the
// actual list comes from the live _Options sheet tab at runtime (see
// resolveOptions() below and Code.gs's getOptionsMap()), so adding,
// removing, or reordering choices is a sheet edit, not a code change.
// These DEFAULT_OPTIONS are the fallback used if the sheet's copy of a key
// is missing/empty (including the very first run, before _Options has
// even been created yet) — the app never breaks over a missing key, it
// just falls back to what it always shipped with. They also double as the
// exact starting rows ensureOptionsTab() seeds the sheet with, so editing
// there means editing real values, not building a list from scratch.
const DEFAULT_OPTIONS = {
  SITES: ['MHF', 'FV', 'Peds', 'GI', 'Specialty Pharmacy', 'All'],
  REQUESTERS: ['Cassandra', 'Tseten', 'Grant', 'Tammy', 'Naveen'],
  ASSIGNEES: ['Naveen', 'Surya', 'Amulya', 'Cassandra', 'Tseten', 'Grant', 'Tammy', 'Lucy', 'Erika'],
  // Deliberately its own key, not ASSIGNEES — Leave is scoped to just this
  // smaller group, while ASSIGNEES (the wider team) still backs "Assigned
  // to" on every other tab and shouldn't shrink along with this one.
  LEAVE_NAMES: ['Amulya Kumar', 'Suryaraj', 'Naveen Raj'],
  'Daily Status.Status': ['Done', 'Pending', 'In Progress', 'Open'],
  'Daily Status.Priority': ['High', 'Medium', 'Low'],
  'Cleanup_Activities.Volume': ['Large', 'Medium', 'Small'],
  'Drupal_Bugs_&_Improvements.Type': ['Bug', 'Fix', 'Suggestion'],
  'Leave.Type': ['Vacation', 'Sick', 'WFH', 'Holiday', 'Other'],
};

// `liveOptions` is whatever came back from GET ?action=options (a plain
// object of key -> array), fetched once at launch — see index.html. Falls
// back to DEFAULT_OPTIONS per-key, not all-or-nothing, so a sheet that
// only defines some lists still gets the rest from the shipped defaults.
function resolveOptions(key, liveOptions) {
  const fromSheet = liveOptions && liveOptions[key];
  return fromSheet && fromSheet.length ? fromSheet : DEFAULT_OPTIONS[key] || [];
}

// Top-level menu the widget shows before the tab list — started life as
// just "log a status update", now covers more than one kind of thing, so
// tabs group under whichever menu they conceptually belong to instead of
// one long flat list. A tab that shows up from the sheet but isn't listed
// under any of these still appears — see groupTabsIntoCategories() below —
// bucketed into an auto-created "More" menu, so a newly added tab is
// never silently unreachable, just uncategorized until it's added here.
const CATEGORIES = {
  'Report Generator': ['Daily Status', 'Adhoc_Mails', 'Cleanup_Activities', 'Drupal_Bugs_&_Improvements'],
  'Team Management': ['Leave'],
};

const FIELD_CONFIG = {
  'Daily Status': {
    Site: { type: 'select', optionsKey: 'SITES' },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    Status: { type: 'select', optionsKey: 'Daily Status.Status' },
    Priority: { type: 'select', optionsKey: 'Daily Status.Priority' },
    'Assigned to': { type: 'multiselect', optionsKey: 'ASSIGNEES' },
  },
  Adhoc_Mails: {
    Requester: { type: 'select', optionsKey: 'REQUESTERS' },
    Site: { type: 'select', optionsKey: 'SITES' },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    'Assigned to': { type: 'multiselect', optionsKey: 'ASSIGNEES' },
  },
  Cleanup_Activities: {
    'Cleanup Number': { type: 'sequence' },
    Volume: { type: 'select', optionsKey: 'Cleanup_Activities.Volume' },
    Requester: { type: 'select', optionsKey: 'REQUESTERS' },
    'Site Impacted': { type: 'select', optionsKey: 'SITES' },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    'Assigned to': { type: 'multiselect', optionsKey: 'ASSIGNEES' },
  },
  'Drupal_Bugs_&_Improvements': {
    Website: { type: 'select', optionsKey: 'SITES' },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    Type: { type: 'select', optionsKey: 'Drupal_Bugs_&_Improvements.Type' },
    'Assigned to': { type: 'multiselect', optionsKey: 'ASSIGNEES' },
  },
  // Auto-created server-side (see ensureLeaveTab() in Code.gs) the first
  // time the widget asks for the tab list — this config just makes that
  // tab's form as polished as the rest instead of falling back to plain
  // text boxes for every field. One person per leave entry, so a plain
  // `select` (not `multiselect` like "Assigned to" elsewhere) fits better.
  Leave: {
    Name: { type: 'select', optionsKey: 'LEAVE_NAMES' },
    Date: { type: 'date' },
    Week: { type: 'week-auto', basedOn: 'Date' },
    Type: { type: 'select', optionsKey: 'Leave.Type' },
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
