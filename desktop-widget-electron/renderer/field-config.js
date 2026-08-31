// Nothing in this file is hardcoded per-organization data any more —
// every list, every field's widget type, and every category is fetched
// from the connected sheet at runtime (see index.html's refreshLiveOptions
// / refreshFieldSchema / refreshCategories and Code.gs's getOptionsMap() /
// getFieldSchemaMap() / getCategoriesMap()). A brand-new organization that
// hasn't configured any of _Options / _FieldSchema / _Categories yet still
// gets a fully working app: every field just falls back to a plain text
// box, and the tab list shows flat with no grouping screen at all — see
// the fallbacks below and in buildFieldInput() (index.html).

// `liveOptions` is whatever came back from GET ?action=options (a plain
// object of key -> array), fetched once at launch — see index.html.
// Returns [] for a key the sheet doesn't define, rather than a shipped
// default — there's no "correct" default site list, requester list, etc.
// for an organization we know nothing about.
function resolveOptions(key, liveOptions) {
  const fromSheet = liveOptions && liveOptions[key];
  return fromSheet && fromSheet.length ? fromSheet : [];
}

function normalizeKey(s) {
  return String(s).trim().toLowerCase();
}

// Buckets the live list of sheet tabs into whatever categories the
// connected sheet's _Categories tab defines, in that tab's own row order,
// skipping a category entirely if none of its tabs currently exist (e.g. a
// tab got renamed). Anything left over — present on the sheet but not
// listed under any category — lands in a "More" bucket appended at the
// end, so a brand-new tab is discoverable instead of just vanishing.
//
// If liveCategories is empty (nothing configured at all — the default for
// a brand-new organization), returns [] on purpose: that's the signal
// index.html uses to skip the category-picker screen entirely and show a
// flat tab list instead, exactly like the app behaved before categories
// existed. A "More" bucket only ever appears alongside *some* real
// configured category, never as the sole group.
function groupTabsIntoCategories(tabs, liveCategories) {
  const categoryNames = Object.keys(liveCategories || {});
  if (categoryNames.length === 0) return [];

  const used = new Set();
  const groups = categoryNames
    .map((name) => {
      const members = liveCategories[name].filter((t) => tabs.indexOf(t) !== -1);
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
// won't always match the _FieldSchema tab's Column value exactly.
// fieldSchema is whatever GET ?action=fieldSchema returned: { Tab: { Column:
// {type, optionsKey, basedOn} } }. A tab/column not present there simply
// has no spec — buildFieldInput() (index.html) falls back to a plain text
// box in that case, which is the correct behavior for an org that hasn't
// configured that field yet.
function fieldSpecFor(tab, column, fieldSchema) {
  const tabConfig = fieldSchema && fieldSchema[tab];
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
