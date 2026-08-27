// Per-tab field behavior: which fields get dropdowns, a date picker, or an
// auto-computed "Week" — everything else falls back to a plain text box.
// Tab names here must match the actual sheet tab names exactly.

const SITES = ['MHF', 'FV', 'Peds', 'GI', 'Specialty Pharmacy', 'All'];
const REQUESTERS = ['Cassandra', 'Tseten', 'Grant', 'Tammy', 'Naveen'];
const ASSIGNEES = ['Naveen', 'Surya', 'Amulya', 'Cassandra', 'Tseten', 'Grant', 'Tammy', 'Lucy', 'Erika'];

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
};

function normalizeKey(s) {
  return String(s).trim().toLowerCase();
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

function computeWeekLabel(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  if (isNaN(d.getTime())) return '';
  const dayIndex = (d.getDay() + 6) % 7; // Monday = 0
  const monday = new Date(d);
  monday.setDate(d.getDate() - dayIndex);
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);
  const fmt = (x) => x.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  return `${fmt(monday)} – ${fmt(friday)}, ${friday.getFullYear()}`;
}
