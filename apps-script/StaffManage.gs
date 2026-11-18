/**
 * Staff Manage tab: the one place staff change who's who and who coaches whom.
 *
 * Kim and the pastors edit the "Staff Manage" tab like a normal spreadsheet, with dropdowns.
 * Every edit is copied into the People and Assignments tabs (which the app reads) within
 * seconds, and written to the "Change Log" tab. When the app itself changes a person or a
 * coach assignment (for example finishing an interview), the matching Staff Manage row is
 * rewritten so the two always agree.
 *
 * Setup (once, from the editor): run setupStaffManage. It adds two People columns, builds the
 * tab from today's data, adds dropdowns, protects the raw tabs, and starts the edit trigger
 * plus a nightly tidy-up that re-sorts the tab.
 */

var SM_TAB = 'Staff Manage';
var SM_LISTS = 'Staff Lists';
var SM_LOG = 'Change Log';
var SM_HEADERS = ['First', 'Last', 'Email', 'Phone', 'Role', 'Coach / Senior Coach', 'Status', 'Ministry',
  'Rooted Season', 'Rooted Facilitator', 'S Level', 'L Level', 'Notes', 'Needs Discord setup', 'Check', 'ID'];
var SM = {}; SM_HEADERS.forEach(function (h, i) { SM[h] = i; });
var SM_ROLES = ['Leader', 'Coach', 'Senior Coach', 'Staff'];
var SM_STATUSES = ['Active', 'Paused', 'Onboarding', 'Not onboarded', 'No longer leading'];
var SM_LEVELS = ['1', '2', '3', '4'];
var SM_ROLE_ORDER = { 'Staff': 0, 'Senior Coach': 1, 'Coach': 2, 'Leader': 3 };

// ---------------------------------------------------------------- helpers
function smToday_() { return Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); }
function smName_(p) { return p ? (p.first + ' ' + p.last).trim() : ''; }
function smNextId_(rows, prefix) {
  var max = 100;
  rows.forEach(function (r) {
    var m = String(r.id).match(new RegExp('^' + prefix + '-(\\d+)$'));
    if (m) max = Math.max(max, Number(m[1]));
  });
  return prefix + '-' + (max + 1);
}
function smMinistries_() {
  var byId = {}, byName = {};
  read_('ministries').forEach(function (m) { byId[m.id] = m.name; byName[String(m.name).toLowerCase()] = m.id; });
  return { byId: byId, byName: byName };
}
/** Coach dropdown labels: "First Last", or "First Last (Ministry)" when two coaches share a name. */
function smCoachLabels_(people) {
  var min = smMinistries_();
  var coaches = people.filter(function (p) { return (p.role === 'Coach' || p.role === 'Senior Coach') && p.status !== 'No longer leading'; });
  var count = {};
  coaches.forEach(function (p) { var n = smName_(p).toLowerCase(); count[n] = (count[n] || 0) + 1; });
  var labelOf = {}, idOf = {};
  coaches.forEach(function (p) {
    var label = smName_(p);
    if (count[label.toLowerCase()] > 1) label += ' (' + (min.byId[p.ministry] || p.id) + ')';
    labelOf[p.id] = label;
    idOf[label.toLowerCase()] = p.id;
  });
  return { labelOf: labelOf, idOf: idOf, list: Object.keys(labelOf).map(function (id) { return labelOf[id]; }).sort() };
}
function smOpenAssignments_() {
  var coachOf = {}, rowOf = {};
  read_('assignments').forEach(function (a) { if (!a.end) { coachOf[a.leader] = a.coach; rowOf[a.leader] = a; } });
  return { coachOf: coachOf, rowOf: rowOf };
}

/** One Staff Manage row (without Check) for a person, from People + Assignments. */
function smRowFromData_(p, coachOf, labels, min) {
  var row = SM_HEADERS.map(function () { return ''; });
  row[SM['First']] = p.first || '';
  row[SM['Last']] = p.last || '';
  row[SM['Email']] = p.email || '';
  row[SM['Phone']] = p.phone ? "'" + p.phone : '';
  row[SM['Role']] = p.role || '';
  var above = p.role === 'Leader' ? coachOf[p.id] : p.role === 'Coach' ? p.seniorCoach : '';
  row[SM['Coach / Senior Coach']] = above ? (labels.labelOf[above] || '') : '';
  row[SM['Status']] = p.status || '';
  row[SM['Ministry']] = min.byId[p.ministry] || '';
  row[SM['Rooted Season']] = p.rootedSeason || '';
  row[SM['Rooted Facilitator']] = p.rootedFac || '';
  row[SM['S Level']] = p.sLevel ? String(p.sLevel) : '';
  row[SM['L Level']] = p.lLevel ? String(p.lLevel) : '';
  row[SM['Notes']] = p.staffNotes || '';
  row[SM['Needs Discord setup']] = p.needsDiscord === true;
  row[SM['ID']] = p.id;
  return row;
}

/** People in display order: staff, then each senior coach's group (senior coach, their direct leaders, each coach and their leaders), then anyone left over. */
function smOrdered_(people, coachOf) {
  var byName = function (a, b) { return smName_(a).localeCompare(smName_(b)); };
  var used = {}, out = [];
  var add = function (p) { if (p && !used[p.id]) { used[p.id] = true; out.push(p); } };
  var leadersOf = function (id) { return people.filter(function (p) { return p.role === 'Leader' && coachOf[p.id] === id; }).sort(byName); };
  people.filter(function (p) { return p.role === 'Staff'; }).sort(byName).forEach(add);
  people.filter(function (p) { return p.role === 'Senior Coach'; }).sort(byName).forEach(function (sc) {
    add(sc);
    leadersOf(sc.id).forEach(add);
    people.filter(function (p) { return p.role === 'Coach' && p.seniorCoach === sc.id; }).sort(byName).forEach(function (c) {
      add(c);
      leadersOf(c.id).forEach(add);
    });
  });
  people.filter(function (p) { return p.role === 'Coach'; }).sort(byName).forEach(function (c) { add(c); leadersOf(c.id).forEach(add); });
  people.slice().sort(byName).forEach(add);
  return out;
}

// ---------------------------------------------------------------- building the tab
/** Rewrites the Staff Manage tab from People + Assignments, sorted. Also refreshes the dropdown lists. */
function smRebuild_() {
  var ss = ss_();
  CACHE_ = {};
  var people = read_('people');
  var coachOf = smOpenAssignments_().coachOf;
  var labels = smCoachLabels_(people);
  var min = smMinistries_();
  smWriteLists_(labels);

  var sh = ss.getSheetByName(SM_TAB) || ss.insertSheet(SM_TAB, 0);
  var rows = smOrdered_(people, coachOf).map(function (p) {
    var r = smRowFromData_(p, coachOf, labels, min);
    r[SM['Check']] = '✓';
    return r;
  });
  sh.clearContents();
  sh.getRange(1, 1, 1, SM_HEADERS.length).setValues([SM_HEADERS]);
  if (rows.length) sh.getRange(2, 1, rows.length, SM_HEADERS.length).setValues(rows);
  smFormat_(sh);
}

function smWriteLists_(labels) {
  var ss = ss_();
  var sh = ss.getSheetByName(SM_LISTS);
  var min = read_('ministries').map(function (m) { return m.name; });
  if (!sh) {
    sh = ss.insertSheet(SM_LISTS);
    var seasons = [];
    var year = new Date().getFullYear();
    for (var y = year + 1; y >= year - 2; y--) ['Fall', 'Spring', 'Winter'].forEach(function (s) { seasons.push(s + ' ' + y); });
    seasons.push('Other');
    sh.getRange(1, 1, 1, 5).setValues([['Rooted Seasons (staff can add more)', 'Coaches (automatic)', 'Roles', 'Statuses', 'Ministries']]);
    sh.getRange(2, 1, seasons.length, 1).setValues(seasons.map(function (s) { return [s]; }));
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var put = function (col, values) {
    sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1).clearContent();
    if (values.length) sh.getRange(2, col, values.length, 1).setValues(values.map(function (v) { return [v]; }));
  };
  put(2, labels.list);
  put(3, SM_ROLES);
  put(4, SM_STATUSES);
  put(5, min);
}

function smFormat_(sh) {
  var lists = ss_().getSheetByName(SM_LISTS);
  var n = Math.max(sh.getMaxRows() - 1, 1);
  var col = function (h) { return sh.getRange(2, SM[h] + 1, n, 1); };
  var listRule = function (range) { return SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(false).build(); };
  var fixed = function (values) { return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build(); };

  col('Role').setDataValidation(fixed(SM_ROLES));
  col('Status').setDataValidation(fixed(SM_STATUSES));
  col('S Level').setDataValidation(fixed(SM_LEVELS));
  col('L Level').setDataValidation(fixed(SM_LEVELS));
  col('Coach / Senior Coach').setDataValidation(listRule(lists.getRange('B2:B')));
  col('Ministry').setDataValidation(listRule(lists.getRange('E2:E')));
  col('Rooted Season').setDataValidation(listRule(lists.getRange('A2:A')));
  col('Needs Discord setup').insertCheckboxes();
  col('Phone').setNumberFormat('@');

  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  sh.getRange(1, 1, 1, SM_HEADERS.length).setFontWeight('bold').setBackground('#e8eef7');
  col('Check').setFontColor('#5f6368');
  col('ID').setFontColor('#9aa0a6');
  sh.getRange(1, SM['Check'] + 1, 1, 2).setBackground('#f1f3f4');
  sh.setColumnWidth(SM['Notes'] + 1, 320);
  sh.setColumnWidth(SM['Coach / Senior Coach'] + 1, 170);
  sh.setColumnWidth(SM['Check'] + 1, 260);
  sh.getRange(1, SM['Check'] + 1).setNote('Filled in automatically: ✓ when a row is saved, or what needs fixing.');
  sh.getRange(1, SM['ID'] + 1).setNote('Filled in automatically. Leave it alone; it links this row to the app.');
  sh.getRange(1, SM['Coach / Senior Coach'] + 1).setNote('Leaders: pick their coach. Coaches: pick their senior coach. Senior coaches and staff: leave blank.');
  sh.getRange(1, SM['Ministry'] + 1).setNote('Filled in from their coach if left blank. Senior coaches need one.');
  sh.getRange(1, SM['Rooted Season'] + 1).setNote('How they joined. If it isn\'t a Rooted season, pick Other and say how in Notes.');
}

// ---------------------------------------------------------------- syncing an edited row
/**
 * Saves one Staff Manage row into People and Assignments. Returns what to show in Check,
 * and the change-log lines. Does nothing for empty rows.
 */
function smSyncRow_(sh, r, by) {
  var v = sh.getRange(r, 1, 1, SM_HEADERS.length).getValues()[0];
  var get = function (h) { var x = v[SM[h]]; return x === true || x === false ? x : String(x === null || x === undefined ? '' : x).trim(); };
  var check = function (msg) { sh.getRange(r, SM['Check'] + 1).setValue(msg); return { msg: msg, log: [] }; };
  var id = get('ID');
  var first = get('First'), last = get('Last'), role = get('Role');
  var filled = SM_HEADERS.some(function (h) { return h !== 'Check' && h !== 'ID' && h !== 'Needs Discord setup' && get(h) !== ''; });
  if (!filled) { if (!id) sh.getRange(r, SM['Check'] + 1).setValue(''); return { msg: '', log: [] }; }
  if (!first || !last || !role) return check('⚠ Needs First, Last and Role');
  if (SM_ROLES.indexOf(role) < 0) return check('⚠ Pick a Role from the list');

  CACHE_ = {};
  var people = read_('people');
  var byId = {}; people.forEach(function (p) { byId[p.id] = p; });
  var before = id ? byId[id] : null;
  if (id && !before) return check('⚠ ID ' + id + ' isn\'t in People. Clear the ID to add them as new.');
  var labels = smCoachLabels_(people);
  var min = smMinistries_();
  var open = smOpenAssignments_();

  // who they report to
  var aboveName = get('Coach / Senior Coach');
  var above = aboveName ? labels.idOf[aboveName.toLowerCase()] : '';
  if (aboveName && !above) return check('⚠ "' + aboveName + '" isn\'t a coach. Pick from the list.');
  if (role === 'Leader' && !above) return check('⚠ Pick their coach');
  if (role === 'Coach') {
    if (!above) return check('⚠ Pick their senior coach');
    if (byId[above].role !== 'Senior Coach') return check('⚠ A coach\'s Coach / Senior Coach must be a senior coach');
  }
  if (above && above === id) return check('⚠ Can\'t coach themselves');

  // someone who coaches leaders can't stop being a coach until their leaders move
  if (before && (before.role === 'Coach' || before.role === 'Senior Coach') && role !== before.role && role !== 'Senior Coach' && role !== 'Coach') {
    var still = Object.keys(open.coachOf).filter(function (l) { return open.coachOf[l] === before.id; }).length;
    if (still) return check('⚠ Still coaches ' + still + (still === 1 ? ' leader' : ' leaders') + '. Move them first, then change the role.');
  }
  if (before && before.role === 'Senior Coach' && role !== 'Senior Coach') {
    var coachesUnder = people.filter(function (p) { return p.seniorCoach === before.id && p.role === 'Coach' && p.status !== 'No longer leading'; }).length;
    if (coachesUnder) return check('⚠ Still over ' + coachesUnder + (coachesUnder === 1 ? ' coach' : ' coaches') + '. Move them first, then change the role.');
  }

  var ministryName = get('Ministry');
  var ministry = ministryName ? min.byName[ministryName.toLowerCase()] : (above ? byId[above].ministry : (before && before.ministry) || '');
  if (ministryName && !ministry) return check('⚠ Pick a Ministry from the list');
  if (role === 'Senior Coach' && !ministry) return check('⚠ Senior coaches need a Ministry');

  var status = get('Status') || 'Active';
  if (SM_STATUSES.indexOf(status) < 0) return check('⚠ Pick a Status from the list');
  var level = function (h) { var x = get(h); if (x === '') return ''; var n = Number(x); return n >= 1 && n <= 4 ? n : NaN; };
  var sLevel = level('S Level'), lLevel = level('L Level');
  if (sLevel !== sLevel || lLevel !== lLevel) return check('⚠ Levels are 1 to 4');

  var email = get('Email');
  if (!before) {
    var same = people.filter(function (p) {
      return (email && String(p.email || '').toLowerCase() === email.toLowerCase()) ||
        (smName_(p).toLowerCase() === (first + ' ' + last).toLowerCase() && p.role === role);
    })[0];
    if (same) return check('⚠ Looks like ' + smName_(same) + ' (' + same.id + ') is already listed. Edit that row instead, or change the name/email.');
  }

  var row = {
    id: id || smNextId_(people, 'P'), first: first, last: last, email: email, phone: get('Phone'), role: role,
    ministry: ministry || '', status: status, seniorCoach: role === 'Coach' ? above : '',
    sLevel: sLevel, lLevel: lLevel, staffNotes: get('Notes'), rootedSeason: get('Rooted Season'),
    rootedFac: get('Rooted Facilitator'), needsDiscord: get('Needs Discord setup') === true,
  };
  if (!before) {
    row.created = smToday_(); row.signInEmail = email; row.rhythmWeeks = 6; row.isInterviewer = false; row.canReview = false;
  }

  // what changed, for the log
  var log = [];
  var who = first + ' ' + last;
  if (!before) log.push('Added ' + who + ' as ' + role + (above ? ' with ' + smName_(byId[above]) : ''));
  else {
    var fields = [['first', 'First'], ['last', 'Last'], ['email', 'Email'], ['phone', 'Phone'], ['role', 'Role'], ['status', 'Status'],
      ['sLevel', 'S Level'], ['lLevel', 'L Level'], ['rootedSeason', 'Rooted Season'], ['rootedFac', 'Rooted Facilitator'], ['staffNotes', 'Notes']];
    fields.forEach(function (f) {
      var a = before[f[0]] === undefined || before[f[0]] === null ? '' : String(before[f[0]]);
      var b = row[f[0]] === undefined || row[f[0]] === null ? '' : String(row[f[0]]);
      if (a !== b) log.push(who + ': ' + f[1] + ' "' + a + '" → "' + b + '"');
    });
    if ((before.ministry || '') !== (row.ministry || '')) log.push(who + ': Ministry "' + (min.byId[before.ministry] || '') + '" → "' + (min.byId[row.ministry] || '') + '"');
    if (role === 'Coach' && (before.seniorCoach || '') !== above) log.push(who + ': senior coach ' + (smName_(byId[before.seniorCoach]) || 'none') + ' → ' + smName_(byId[above]));
  }

  upsertRow_('people', row);

  // a coach or senior coach renamed: update the dropdown text everywhere it's used
  if (before && (role === 'Coach' || role === 'Senior Coach') && labels.labelOf[before.id] && smName_(before) !== first + ' ' + last) {
    var oldLabel = labels.labelOf[before.id];
    CACHE_ = {};
    var newLabels = smCoachLabels_(read_('people'));
    smWriteLists_(newLabels);
    var col = sh.getRange(2, SM['Coach / Senior Coach'] + 1, Math.max(sh.getLastRow() - 1, 1), 1);
    col.setValues(col.getValues().map(function (x) { return [String(x[0]) === oldLabel ? newLabels.labelOf[before.id] : x[0]]; }));
  }

  // leaders: keep exactly one open assignment, pointing at the chosen coach
  var current = open.rowOf[row.id];
  if (role === 'Leader') {
    if (!current || current.coach !== above) {
      var assignments = read_('assignments');
      if (current) upsertRow_('assignments', { id: current.id, end: smToday_() });
      upsertRow_('assignments', {
        id: smNextId_(assignments, 'A'), leader: row.id, coach: above, start: smToday_(), temporary: false,
        reason: current ? 'Moved in Staff Manage' : 'Added in Staff Manage', movedBy: by.id || '', externalUpdated: false,
      });
      if (before) log.push(who + ': coach ' + (current ? smName_(byId[current.coach]) : 'none') + ' → ' + smName_(byId[above]));
    }
  } else if (current) {
    upsertRow_('assignments', { id: current.id, end: smToday_() }); // no longer a leader
  }

  // coach list changes when a coach is added, renamed, or changes role/status
  if (!before || role === 'Coach' || role === 'Senior Coach' || (before && (before.role === 'Coach' || before.role === 'Senior Coach'))) {
    CACHE_ = {};
    smWriteLists_(smCoachLabels_(read_('people')));
  }

  sh.getRange(r, SM['ID'] + 1).setValue(row.id);
  if (!ministryName && ministry) sh.getRange(r, SM['Ministry'] + 1).setValue(min.byId[ministry] || '');
  if (!get('Status')) sh.getRange(r, SM['Status'] + 1).setValue(status);
  return { msg: check('✓ Saved ' + smToday_()).msg, log: log };
}

function smLog_(by, lines) {
  if (!lines.length) return;
  var ss = ss_();
  var sh = ss.getSheetByName(SM_LOG);
  if (!sh) {
    sh = ss.insertSheet(SM_LOG);
    sh.appendRow(['When', 'Who', 'What changed']);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var now = new Date();
  sh.getRange(sh.getLastRow() + 1, 1, lines.length, 3).setValues(lines.map(function (l) { return [now, by.name || 'Staff', l]; }));
}

/** Runs on every edit of the spreadsheet (installed by setupStaffManage). */
function smOnEdit(e) {
  if (!e || !e.range || e.range.getSheet().getName() !== SM_TAB) return;
  var sh = e.range.getSheet();
  var top = Math.max(e.range.getRow(), 2), bottom = e.range.getLastRow();
  if (bottom < 2) return;
  if (e.range.getColumn() > SM['Check']) return; // only the automatic Check / ID columns were touched
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { sh.getRange(top, SM['Check'] + 1).setValue('⚠ Busy, edit this row again in a moment'); return; }
  try {
    var by = smEditor_(e);
    var lines = [];
    for (var r = top; r <= bottom; r++) {
      try { lines = lines.concat(smSyncRow_(sh, r, by).log); }
      catch (err) { sh.getRange(r, SM['Check'] + 1).setValue('⚠ ' + String(err && err.message || err).slice(0, 120)); }
    }
    smLog_(by, lines);
  } finally {
    lock.releaseLock();
  }
}

function smEditor_(e) {
  var email = '';
  try { email = (e.user && e.user.getEmail && e.user.getEmail()) || Session.getActiveUser().getEmail() || ''; } catch (err) {}
  email = String(email).toLowerCase();
  var p = email && read_('people').filter(function (x) {
    return String(x.email || '').toLowerCase() === email || String(x.signInEmail || '').toLowerCase() === email;
  })[0];
  return { id: p ? p.id : '', name: p ? smName_(p) : (email || 'Staff') };
}

// ---------------------------------------------------------------- keeping the tab in step with the app
/**
 * After the app changes people or assignments, rewrite those people's rows in Staff Manage
 * (or add them at the bottom). Called from save_ in Code.gs.
 */
function smReflect_(personIds) {
  if (!personIds || !personIds.length) return;
  var sh = ss_().getSheetByName(SM_TAB);
  if (!sh) return;
  CACHE_ = {};
  var people = read_('people');
  var byId = {}; people.forEach(function (p) { byId[p.id] = p; });
  var coachOf = smOpenAssignments_().coachOf;
  var labels = smCoachLabels_(people);
  var min = smMinistries_();
  var last = sh.getLastRow();
  var ids = last > 1 ? sh.getRange(2, SM['ID'] + 1, last - 1, 1).getValues().map(function (x) { return String(x[0]).trim(); }) : [];
  var touchedCoaches = false;
  personIds.forEach(function (pid) {
    var p = byId[pid];
    if (!p) return;
    if (p.role === 'Coach' || p.role === 'Senior Coach') touchedCoaches = true;
    var row = smRowFromData_(p, coachOf, labels, min);
    row[SM['Check']] = '✓ Updated from the app ' + smToday_();
    var at = ids.indexOf(pid);
    if (at >= 0) sh.getRange(at + 2, 1, 1, row.length).setValues([row]);
    else { sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]); ids.push(pid); }
  });
  if (touchedCoaches) smWriteLists_(labels);
}

// ---------------------------------------------------------------- run from the editor
/**
 * Run once. Adds the People columns this needs, builds the Staff Manage tab from today's data,
 * protects the raw tabs so only the owner can change them, and starts the triggers.
 */
function setupStaffManage() {
  var ss = ss_();
  var people = ss.getSheetByName('People');
  var headers = people.getRange(1, 1, 1, people.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  ['Rooted Season', 'Rooted Facilitator', 'Needs Discord Setup'].forEach(function (h) {
    if (headers.indexOf(h) < 0) { people.getRange(1, people.getLastColumn() + 1).setValue(h); headers.push(h); }
  });
  CACHE_ = {};
  smRebuild_();
  smLog_({ name: 'Setup' }, ['Staff Manage tab built from People and Assignments']);

  // raw tabs: only the owner can edit; staff work in Staff Manage, Staff Lists and the app
  var me = Session.getEffectiveUser();
  ss.getSheets().forEach(function (s) {
    var name = s.getName();
    if (name === SM_TAB || name === SM_LISTS) return;
    var existing = s.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
    var p = existing || s.protect();
    p.setDescription('Managed by The Huddle. Make changes in the Staff Manage tab.');
    p.addEditor(me);
    p.removeEditors(p.getEditors().filter(function (u) { return u.getEmail() !== me.getEmail(); }));
    if (p.canDomainEdit()) p.setDomainEdit(false);
  });
  var sm = ss.getSheetByName(SM_TAB);
  sm.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  var auto = sm.getRange(1, SM['Check'] + 1, sm.getMaxRows(), 2).protect().setDescription('Filled in automatically').setWarningOnly(true);
  var lists = ss.getSheetByName(SM_LISTS);
  lists.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { p.remove(); });
  lists.getRange('B:E').protect().setDescription('Filled in automatically. Add Rooted seasons in column A.').setWarningOnly(true);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'smOnEdit' || f === 'tidyStaffManage') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('smOnEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('tidyStaffManage').timeBased().atHour(2).everyDays(1).create();
  Logger.log('Staff Manage is ready: ' + (sm.getLastRow() - 1) + ' people. Edits now sync automatically.');
}

/** Re-sorts the tab into groups (senior coach → coaches → leaders). Runs nightly; safe to run any time. */
function tidyStaffManage() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    var sh = ss_().getSheetByName(SM_TAB);
    if (!sh) return;
    // keep rows that still need fixing (a ⚠ in Check, or no ID yet) exactly as typed, so nobody's work is lost
    var last = sh.getLastRow();
    var pending = last > 1 ? sh.getRange(2, 1, last - 1, SM_HEADERS.length).getValues().filter(function (r) {
      var filled = r.slice(0, SM['Check']).some(function (x) { return x !== '' && x !== false; });
      return filled && (!String(r[SM['ID']]).trim() || /^⚠/.test(String(r[SM['Check']])));
    }) : [];
    smRebuild_();
    var ids = sh.getLastRow() > 1 ? sh.getRange(2, SM['ID'] + 1, sh.getLastRow() - 1, 1).getValues().map(function (x) { return String(x[0]).trim(); }) : [];
    var extra = [];
    pending.forEach(function (r) {
      var at = ids.indexOf(String(r[SM['ID']]).trim());
      if (String(r[SM['ID']]).trim() && at >= 0) sh.getRange(at + 2, 1, 1, SM_HEADERS.length).setValues([r]);
      else extra.push(r);
    });
    if (extra.length) sh.getRange(sh.getLastRow() + 1, 1, extra.length, SM_HEADERS.length).setValues(extra);
  } finally {
    lock.releaseLock();
  }
}
