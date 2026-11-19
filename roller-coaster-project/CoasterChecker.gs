/**
 * ROLLER COASTER POLYNOMIAL CHECKER  (Google Apps Script)
 * ------------------------------------------------------------------
 * What it does
 *   1. Builds a Google Form for students (one click from the menu).
 *   2. Every time a student submits, it reads their FACTORED (intercept)
 *      form, expands it to STANDARD form, and checks every requirement
 *      of the project. Results land on the "Results" tab, color coded.
 *   3. Gives you an answer key for each student: turning points,
 *      intervals of increase/decrease, range, end behavior.
 *
 * Setup (about 3 minutes, done once)
 *   1. Create a new, blank Google Sheet.
 *   2. Extensions > Apps Script. Delete what is there, paste this whole
 *      file, click Save (disk icon).
 *   3. Go back to the Sheet and refresh the page. A menu called
 *      "Coaster Checker" appears. Choose  Coaster Checker > 1. Create student form.
 *   4. Google will ask you to authorize the script (it's your own script).
 *      Click through: Advanced > Go to project > Allow.
 *   5. The form link appears in a pop-up and on the "Settings" tab.
 *      Post that link in Google Classroom.
 *
 * You can change the requirements (ride length, number of turning points,
 * tolerances...) on the "Settings" tab at any time, then use
 * Coaster Checker > Re-check all submissions.
 * ------------------------------------------------------------------
 */

// ===================== FORM QUESTION TITLES =========================
// The checker finds answers by these titles. If you rename a question in
// the Form, rename it here too.
var Q = {
  name:      'Your name (first and last)',
  period:    'Class period',
  coaster:   'Name of your roller coaster',
  factored:  'Factored (intercept) form of your polynomial',
  pointX:    'Point used to find a: TIME (seconds)',
  pointY:    'Point used to find a: HEIGHT (feet)',
  standard:  'Standard form of your polynomial',
  rideEnd:   'Ride end time (seconds)',
  endLeft:   'End behavior: as x → −∞, y → ?',
  endRight:  'End behavior: as x → +∞, y → ?',
  turning:   'Turning point times (seconds)',
  rangeLow:  'Range: LOWEST height during the ride (feet)',
  rangeHigh: 'Range: HIGHEST height during the ride (feet)',
  desmos:    'Desmos graph link'
};

var DEFAULT_SETTINGS = [
  // key, value, explanation
  ['Minimum ride length (seconds)', 120, 'Ride end time must be at least this.'],
  ['Minimum turning points', 3, 'Counted only during the ride (0 to end time).'],
  ['Require a repeated zero (multiplicity 2 or more)', 'NO', 'YES or NO'],
  ['Highest realistic height (feet)', 500, 'Taller than this gets a warning (tallest coaster ever was 456 ft).'],
  ['Deepest realistic tunnel (feet)', -150, 'Deeper than this gets a warning.'],
  ['a-value tolerance (%)', 2, 'How close the student a-value must be to the exact value.'],
  ['Standard form tolerance (%)', 1, 'How close each standard-form coefficient must be to the expansion of the student\'s factored form.'],
  ['Turning point time tolerance (seconds)', 1, 'Student turning point times may be off by this much.'],
  ['Range tolerance (feet)', 2, 'Student lowest/highest heights may be off by this much (or 2% of the height span, whichever is bigger).'],
  ['Email students their check results', 'NO', 'YES = students get an email after submitting showing which checks passed (it never shows them the answers).']
];

// ===================== MENU / SETUP =================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Coaster Checker')
    .addItem('1. Create student form', 'setupForm')
    .addItem('Re-check all submissions', 'recheckAll')
    .addSeparator()
    .addItem('Show form link', 'showFormLink')
    .addItem('Help', 'showHelp')
    .addToUi();
}

function setupForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ensureSettingsSheet_(ss);

  var existing = getSetting_('__formUrl', '');
  if (existing) {
    var r = ui.alert('A form already exists',
      'This spreadsheet already has a student form:\n' + existing +
      '\n\nCreate a brand-new form anyway?', ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
  }

  var form = FormApp.create('Roller Coaster Polynomial – Coaster Check');
  form.setDescription(
    'Submit your roller coaster polynomial here BEFORE you turn in your project.\n\n' +
    'Typing tips:\n' +
    '•  Use ^ for exponents:  (x - 40)^2\n' +
    '•  Very small numbers: type them in full (-0.0000000508) or with E notation (-5.08E-8). ' +
    'Desmos style -5.08×10^-8 also works.\n' +
    '•  Copying straight out of Desmos is fine.\n\n' +
    'You may submit more than once. Your teacher sees your most recent submission.');
  try { form.setCollectEmail(true); } catch (err) { console.warn('Could not turn on email collection: ' + err); }
  form.setProgressBar(false);
  form.setConfirmationMessage('Thanks! Your coaster has been submitted to your teacher.');

  var numberCheck = FormApp.createTextValidation()
    .setHelpText('Please enter a number (digits only, no units).')
    .requireNumber().build();

  form.addTextItem().setTitle(Q.name).setRequired(true);
  form.addTextItem().setTitle(Q.period).setRequired(true);
  form.addTextItem().setTitle(Q.coaster).setRequired(false);

  form.addSectionHeaderItem().setTitle('Your polynomial');
  form.addParagraphTextItem().setTitle(Q.factored).setRequired(true)
    .setHelpText('Include your a-value. Example:  y = -0.0000000508(x - 20)(x - 55)^2(x - 100)(x - 130)');
  form.addTextItem().setTitle(Q.pointX).setRequired(true).setValidation(numberCheck)
    .setHelpText('The point you plugged in to solve for a. Usually your platform: time = 0.');
  form.addTextItem().setTitle(Q.pointY).setRequired(true).setValidation(numberCheck)
    .setHelpText('Usually your starting platform height.');
  form.addParagraphTextItem().setTitle(Q.standard).setRequired(true)
    .setHelpText('Your final answer after multiplying everything out. Example:  y = -5.08E-8x^5 + 0.0000183x^4 - ...');

  form.addSectionHeaderItem().setTitle('Your answers');
  form.addTextItem().setTitle(Q.rideEnd).setRequired(true).setValidation(numberCheck)
    .setHelpText('The time (in seconds) when your ride ends. Your domain is [0, this number].');
  form.addMultipleChoiceItem().setTitle(Q.endLeft).setRequired(true).setChoiceValues(['+∞ (rises)', '−∞ (falls)']);
  form.addMultipleChoiceItem().setTitle(Q.endRight).setRequired(true).setChoiceValues(['+∞ (rises)', '−∞ (falls)']);
  form.addTextItem().setTitle(Q.turning).setRequired(true)
    .setHelpText('The x-values of ALL your maximums and minimums during the ride, separated by commas. Example: 8.2, 37.5, 79.1, 115.4');
  form.addTextItem().setTitle(Q.rangeLow).setRequired(true).setValidation(numberCheck)
    .setHelpText('Lowest point of the ride (the bottom of your tunnel is negative).');
  form.addTextItem().setTitle(Q.rangeHigh).setRequired(true).setValidation(numberCheck);
  form.addTextItem().setTitle(Q.desmos).setRequired(false)
    .setHelpText('Optional: in Desmos click Share > copy link and paste it here.');

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();

  // Re-install the submit trigger (only one).
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onCoasterSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onCoasterSubmit').forSpreadsheet(ss).onFormSubmit().create();

  setSetting_('__formUrl', form.getPublishedUrl());
  setSetting_('__formEditUrl', form.getEditUrl());
  ensureResultsSheet_(ss);

  ui.alert('Form created!',
    'Student link (post this in Classroom):\n' + form.getPublishedUrl() +
    '\n\nEdit the form here:\n' + form.getEditUrl() +
    '\n\nBoth links are also saved at the bottom of the Settings tab.', ui.ButtonSet.OK);
}

function showFormLink() {
  var url = getSetting_('__formUrl', '');
  SpreadsheetApp.getUi().alert(url ? 'Student form link:\n' + url : 'No form yet. Use Coaster Checker > 1. Create student form.');
}

function showHelp() {
  SpreadsheetApp.getUi().alert('Coaster Checker',
    'RESULTS tab: one row per submission.\n' +
    '  ✅ = requirement met   ❌ = needs fixing   ⚠️ = check by hand\n' +
    '  Hover over (or click) a ✅/❌ cell to read the details.\n' +
    '  The gray "answer key" columns are computed from the student\'s own factored form.\n\n' +
    'SETTINGS tab: change requirements and tolerances, then choose\n' +
    '  Coaster Checker > Re-check all submissions.\n\n' +
    'FIXING A TYPO: edit the student\'s answer on the Form Responses tab, then Re-check all.\n\n' +
    'The "Teacher notes" column on the Results tab is kept when you re-check.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ===================== TRIGGER ======================================
function onCoasterSubmit(e) {
  var results = recheckAll(true);
  try {
    if (String(getSetting_(DEFAULT_SETTINGS[9][0], 'NO')).toUpperCase().indexOf('Y') === 0 && e && e.range) {
      var rowIdx = e.range.getRow() - 2; // 0-based data index
      var r = results[rowIdx];
      if (r && r.email) emailStudent_(r);
    }
  } catch (err) { console.error(err); }
}

// ===================== RE-CHECK ALL =================================
var CHECKS = [
  ['readable', 'Equation readable'],
  ['aValue',   'a-value correct'],
  ['standard', 'Standard form matches factored form'],
  ['length',   'Ride long enough'],
  ['turns',    'Enough turning points'],
  ['tunnel',   'Has a tunnel (y < 0)'],
  ['platform', 'Starts on a platform above ground'],
  ['realistic','Realistic heights'],
  ['repeated', 'Repeated zero'],
  ['endBeh',   'End behavior answer'],
  ['turnAns',  'Turning point answer'],
  ['rangeAns', 'Range answer']
];

function recheckAll(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet_(ss);
  var settings = readSettings_();
  var resp = findResponsesSheet_(ss);
  var out = ensureResultsSheet_(ss);
  if (!resp) {
    if (silent !== true) SpreadsheetApp.getUi().alert('No form responses found yet.');
    return [];
  }

  var data = resp.getDataRange().getValues();
  var head = data[0].map(function (h) { return String(h).trim(); });
  function col(title) {
    for (var i = 0; i < head.length; i++) if (head[i] === title) return i;
    for (i = 0; i < head.length; i++) if (head[i].indexOf(title.slice(0, 25)) === 0) return i;
    return -1;
  }
  var idx = {};
  Object.keys(Q).forEach(function (k) { idx[k] = col(Q[k]); });
  idx.email = col('Email Address');
  idx.time = col('Timestamp');

  // keep teacher notes across re-checks
  var headers = resultHeaders_();
  var notesCol = headers.indexOf('Teacher notes');
  var oldNotes = {};
  if (out.getLastRow() > 1) {
    var old = out.getRange(2, 1, out.getLastRow() - 1, headers.length).getValues();
    old.forEach(function (r) { if (r[notesCol]) oldNotes[keyOf_(r[0], r[1], r[2])] = r[notesCol]; });
  }

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row.join('') === '') continue;
    var get = function (k) { return idx[k] >= 0 ? row[idx[k]] : ''; };
    var sub = {
      time: get('time'), email: get('email'), name: get('name'), period: get('period'),
      coaster: get('coaster'), factored: get('factored'), pointX: get('pointX'), pointY: get('pointY'),
      standard: get('standard'), rideEnd: get('rideEnd'), endLeft: get('endLeft'), endRight: get('endRight'),
      turning: get('turning'), rangeLow: get('rangeLow'), rangeHigh: get('rangeHigh'), desmos: get('desmos')
    };
    var res = checkSubmission(sub, settings);
    res.email = sub.email; res.name = sub.name; res.sub = sub;
    results.push(res);
  }

  // mark latest submission per student
  var latest = {};
  results.forEach(function (r, j) {
    var who = String(r.sub.email || r.sub.name).toLowerCase().trim();
    latest[who] = j;
  });

  var rows = [], notes = [], bgs = [];
  results.forEach(function (r, j) {
    var s = r.sub;
    var who = String(s.email || s.name).toLowerCase().trim();
    var line = [s.time, s.name, s.period, latest[who] === j ? '✔' : 'older', s.coaster,
      r.passCount + ' / ' + r.applicable, r.summary];
    var nl = ['', '', '', '', '', '', r.feedback.join('\n')];
    var bg = [null, null, null, latest[who] === j ? null : '#eeeeee', null, null,
      r.passCount === r.applicable ? '#d9ead3' : '#fce5cd'];
    CHECKS.forEach(function (c) {
      var ck = r.checks[c[0]];
      line.push(ck.status === 'pass' ? '✅' : ck.status === 'fail' ? '❌' : ck.status === 'warn' ? '⚠️' : '—');
      nl.push(ck.detail || '');
      bg.push(ck.status === 'pass' ? '#d9ead3' : ck.status === 'fail' ? '#f4cccc' : ck.status === 'warn' ? '#fff2cc' : '#f3f3f3');
    });
    var k = r.key || {};
    line.push(s.factored, k.zerosText || '', k.degree || '', k.aStudent === undefined ? '' : fmtNum(k.aStudent),
      k.aExact === undefined ? '' : fmtNum(k.aExact), k.standardText || '', s.standard,
      k.endBehaviorText || '', k.turningText || '', k.increasingText || '', k.decreasingText || '',
      k.rangeText || '', k.domainText || '', k.endHeightText || '', s.desmos,
      oldNotes[keyOf_(s.time, s.name, s.period)] || '');
    for (var z = 0; z < 16; z++) { nl.push(''); bg.push(z < 14 && z !== 0 && z !== 6 ? '#f3f3f3' : null); }
    rows.push(line); notes.push(nl); bgs.push(bg);
  });

  var last = out.getLastRow();
  if (last > 1) out.getRange(2, 1, last - 1, headers.length).clearContent().clearNote().setBackground(null);
  if (rows.length) {
    var rng = out.getRange(2, 1, rows.length, headers.length);
    rng.setValues(rows); rng.setNotes(notes); rng.setBackgrounds(bgs);
    out.getRange(2, 1, rows.length, 1).setNumberFormat('m/d h:mm am/pm');
    out.getRange(2, 8, rows.length, CHECKS.length).setHorizontalAlignment('center');
  }
  if (silent !== true) ss.toast(rows.length + ' submission(s) checked.', 'Coaster Checker', 5);
  return results;
}

function keyOf_(t, n, p) {
  var ts = (t instanceof Date) ? t.getTime() : String(t);
  return ts + '|' + n + '|' + p;
}

function resultHeaders_() {
  var h = ['Submitted', 'Name', 'Period', 'Latest?', 'Coaster', 'Checks passed', 'Summary'];
  CHECKS.forEach(function (c) { h.push(c[1]); });
  return h.concat(['Factored form (as typed)', 'Zeros (multiplicity)', 'Degree', 'Student a', 'Exact a from their point',
    'CORRECT standard form (from their factored form)', 'Student standard form (as typed)',
    'KEY: End behavior', 'KEY: Turning points (time, height)', 'KEY: Increasing on', 'KEY: Decreasing on',
    'KEY: Range', 'KEY: Domain', 'KEY: Height at end of ride', 'Desmos link', 'Teacher notes']);
}

function ensureResultsSheet_(ss) {
  var sh = ss.getSheetByName('Results');
  var headers = resultHeaders_();
  if (!sh) { sh = ss.insertSheet('Results', 0); }
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold')
    .setBackground('#1f4e79').setFontColor('#ffffff').setWrap(true).setVerticalAlignment('middle');
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  sh.setRowHeight(1, 60);
  sh.setColumnWidth(7, 260);
  for (var c = 8; c < 8 + CHECKS.length; c++) sh.setColumnWidth(c, 95);
  sh.setColumnWidths(8 + CHECKS.length, 1, 280);
  sh.setColumnWidth(13 + CHECKS.length, 360);
  sh.setColumnWidth(14 + CHECKS.length, 300);
  return sh;
}

function findResponsesSheet_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) if (sheets[i].getFormUrl()) return sheets[i];
  for (i = 0; i < sheets.length; i++) if (/^Form Responses/.test(sheets[i].getName())) return sheets[i];
  return null;
}

// ===================== SETTINGS =====================================
function ensureSettingsSheet_(ss) {
  var sh = ss.getSheetByName('Settings');
  if (sh) return sh;
  sh = ss.insertSheet('Settings');
  sh.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'What it means']]).setFontWeight('bold')
    .setBackground('#1f4e79').setFontColor('#ffffff');
  sh.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  sh.getRange(2, 2, DEFAULT_SETTINGS.length, 1).setBackground('#fff2cc').setFontWeight('bold');
  sh.setColumnWidth(1, 330); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 560);
  sh.setFrozenRows(1);
  return sh;
}

function readSettings_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var map = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { map[String(r[0]).trim()] = r[1]; });
  }
  function num(i) { var v = parseNumber(map[DEFAULT_SETTINGS[i][0]]); return isNaN(v) ? DEFAULT_SETTINGS[i][1] : v; }
  function yes(i) { var v = map[DEFAULT_SETTINGS[i][0]]; if (v === undefined || v === '') v = DEFAULT_SETTINGS[i][1]; return String(v).trim().toUpperCase().charAt(0) === 'Y'; }
  return {
    minLength: num(0), minTurns: num(1), requireRepeated: yes(2), maxHeight: num(3), minDepth: num(4),
    aTol: num(5) / 100, coefTol: num(6) / 100, timeTol: num(7), heightTol: num(8)
  };
}

function getSetting_(key, dflt) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sh) return dflt;
  var v = sh.getDataRange().getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === key) return v[i][1];
  return dflt;
}

function setSetting_(key, value) {
  var sh = ensureSettingsSheet_(SpreadsheetApp.getActiveSpreadsheet());
  var v = sh.getDataRange().getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  sh.appendRow([key, value, key === '__formUrl' ? 'Student form link (share this)' : 'Form edit link (teacher only)']);
}

// ===================== EMAIL ========================================
function emailStudent_(r) {
  var lines = ['Hi ' + (r.name || '') + ',', '', 'Here is the automatic check of your roller coaster polynomial.',
    'Checks passed: ' + r.passCount + ' of ' + r.applicable, ''];
  CHECKS.forEach(function (c) {
    var ck = r.checks[c[0]];
    if (ck.status === 'na') return;
    lines.push((ck.status === 'pass' ? '✅ ' : ck.status === 'fail' ? '❌ ' : '⚠️ ') + c[1] +
      (ck.status === 'pass' ? '' : ' — ' + (ck.hint || ck.detail)));
  });
  lines.push('', 'You can fix your work and submit the form again as many times as you need.');
  MailApp.sendEmail(r.email, 'Roller Coaster Polynomial – your check results', lines.join('\n'));
}

// ===================== CUSTOM FUNCTION ==============================
/**
 * Expands a factored-form polynomial into standard form.
 * Example: =COASTER_STANDARD("y=-2(x-1)^2(x+3)")
 * @param {string} factored The factored (intercept) form.
 * @return {string} Standard form.
 * @customfunction
 */
function COASTER_STANDARD(factored) {
  try { return polyToString(parseFactored(String(factored)).coeffs); }
  catch (e) { return 'Could not read: ' + e.message; }
}

// =====================================================================
// =====================  MATH CORE (no Google services)  ==============
// =====================================================================

function normalizeExpr(s) {
  s = String(s === undefined || s === null ? '' : s);
  var sup = { '⁰': '0', '¹': '1', '²': '2', '³': '3', '⁴': '4', '⁵': '5', '⁶': '6', '⁷': '7', '⁸': '8', '⁹': '9', '⁻': '-' };
  s = s.replace(/[⁰¹²³⁴⁵⁶⁷⁸⁹⁻]+/g, function (m) { return '^' + m.split('').map(function (ch) { return sup[ch]; }).join(''); });
  s = s.replace(/[−–—‒]/g, '-')     // minus signs / dashes
       .replace(/[×⋅·∙]/g, '*')     // × ⋅ ·
       .replace(/[\[{]/g, '(').replace(/[\]}]/g, ')')
       .replace(/\\left|\\right|\\cdot|\\times/g, function (m) { return (m === '\\cdot' || m === '\\times') ? '*' : ''; })
       .replace(/\s+/g, '')
       .replace(/,(?=\d{3}(\D|$))/g, '')                // 1,250 -> 1250
       .toLowerCase();
  if (s.indexOf('=') >= 0) s = s.slice(s.lastIndexOf('=') + 1);
  // doubled signs: x--10 -> x+10, x+-10 -> x-10
  for (var g = 0; g < 3; g++) s = s.replace(/--/g, '+').replace(/\+-|-\+/g, '-').replace(/\+\+/g, '+');
  // scientific notation: 5.08*10^-8, 5.08x10^(-8), 5.08×10^{-8}
  s = s.replace(/(\d\.?)(?:\*|x)10\^\(?([+-]?\d+)\)?/g, '$1e$2');
  // a variable other than x (e.g. t)
  if (s.indexOf('x') < 0 && /t/.test(s.replace(/e[+-]?\d/g, ''))) s = s.replace(/t/g, 'x');
  // a bare 10^-8 as a coefficient
  s = s.replace(/(^|[+\-(*])10\^\(?([+-]?\d+)\)?/g, '$11e$2');
  s = s.replace(/\^\((\d+)\)/g, '^$1');                 // ^(2) -> ^2
  s = s.replace(/\(([+-]?\d*\.?\d+(?:e[+-]?\d+)?)\)/g, '$1'); // (-5.08e-8) -> -5.08e-8
  return s;
}

var NUM_RE = '(?:\\d+\\.?\\d*|\\.\\d+)(?:e[+-]?\\d+)?';

function parseNumber(v) {
  if (typeof v === 'number') return v;
  var s = normalizeExpr(v).replace(/[a-df-z]+$/, '');
  if (!new RegExp('^[+-]?' + NUM_RE + '$').test(s)) return NaN;
  return parseFloat(s);
}

function polyMul(p, q) {
  var r = [];
  for (var i = 0; i < p.length + q.length - 1; i++) r.push(0);
  for (i = 0; i < p.length; i++) for (var j = 0; j < q.length; j++) r[i + j] += p[i] * q[j];
  return r;
}

/** Parses a polynomial written in factored form. Returns {a, zeros:[{r,m}], coeffs(low->high)} */
function parseFactored(input) {
  var s = normalizeExpr(input);
  if (!s) throw new Error('blank');
  if (s.indexOf('x') < 0) throw new Error('no x in the equation');
  var pos = 0, a = 1;
  var m = new RegExp('^([+-]?)(' + NUM_RE + ')?\\*?').exec(s);
  if (m && m[0]) {
    var sign = m[1] === '-' ? -1 : 1;
    a = sign * (m[2] ? parseFloat(m[2]) : 1);
    pos = m[0].length;
  }
  var coeffs = [a], zeros = [], leading = a;
  var factorRe = new RegExp('^(?:\\(([^()]*)\\)|x)(?:\\^(\\d+))?\\*?');
  while (pos < s.length) {
    var rest = s.slice(pos);
    var f = factorRe.exec(rest);
    if (!f) {
      var tail = new RegExp('^\\*?(' + NUM_RE + ')\\*?').exec(rest); // a coefficient written at the end/middle
      if (tail && tail[0]) { var k = parseFloat(tail[1]); coeffs = coeffs.map(function (c) { return c * k; }); a *= k; leading *= k; pos += tail[0].length; continue; }
      throw new Error('could not read "' + rest.slice(0, 15) + '"');
    }
    var mult = f[2] ? parseInt(f[2], 10) : 1;
    var lin;
    if (f[1] === undefined) lin = { p: 1, q: 0 };
    else lin = parseLinear(f[1]);
    if (!lin) throw new Error('could not read the factor (' + f[1] + ')');
    var root = -lin.q / lin.p;
    // (px+q) = p(x - root)
    a *= Math.pow(lin.p, mult); leading *= Math.pow(lin.p, mult);
    for (var t = 0; t < mult; t++) coeffs = polyMul(coeffs, [lin.q, lin.p]);
    var found = false;
    zeros.forEach(function (z) { if (Math.abs(z.r - root) < 1e-9) { z.m += mult; found = true; } });
    if (!found) zeros.push({ r: root, m: mult });
    pos += f[0].length;
  }
  if (!zeros.length) throw new Error('no factors like (x - 20) found');
  zeros.sort(function (u, v) { return u.r - v.r; });
  return { a: a, zeros: zeros, coeffs: coeffs };
}

function parseLinear(str) {
  // accepts x-20, x+5, 20-x, 2x-40, -x+3, x
  var t = parseStandard(str, true);
  if (!t || t.length > 2 || !t[1]) return null;
  return { p: t[1], q: t[0] || 0 };
}

/** Parses standard form into coefficient array (low -> high). */
function parseStandard(input, quiet) {
  var s = normalizeExpr(input);
  if (!s) { if (quiet) return null; throw new Error('blank'); }
  // split into terms at + or - that are not exponent signs
  var terms = [], cur = '', depth = 0;
  for (var i = 0; i < s.length; i++) {
    var ch = s.charAt(i), prev = i > 0 ? s.charAt(i - 1) : '';
    if (ch === '(') depth++;
    if (ch === ')') depth--;
    if ((ch === '+' || ch === '-') && depth === 0 && cur !== '' && prev !== 'e' && prev !== '^' && prev !== '*') {
      terms.push(cur); cur = ch;
    } else cur += ch;
  }
  if (cur) terms.push(cur);
  var coeffs = [];
  var termRe = new RegExp('^([+-]?)(' + NUM_RE + ')?\\*?(x(?:\\^(\\d+))?)?$');
  for (i = 0; i < terms.length; i++) {
    var term = terms[i].replace(/^([+-]?)\((.*)\)$/, '$1$2');
    var m = termRe.exec(term);
    if (!m) {
      var alt = new RegExp('^([+-]?)x(?:\\^(\\d+))?\\*(' + NUM_RE + ')$').exec(term); // x^2*3
      if (alt) m = [term, alt[1], alt[3], 'x', alt[2]];
    }
    if (!m || (!m[2] && !m[3])) { if (quiet) return null; throw new Error('could not read the term "' + terms[i] + '"'); }
    var c = (m[1] === '-' ? -1 : 1) * (m[2] ? parseFloat(m[2]) : 1);
    var deg = m[3] ? (m[4] ? parseInt(m[4], 10) : 1) : 0;
    while (coeffs.length <= deg) coeffs.push(0);
    coeffs[deg] += c;
  }
  while (coeffs.length > 1 && coeffs[coeffs.length - 1] === 0) coeffs.pop();
  return coeffs;
}

function evalPoly(c, x) { var y = 0; for (var i = c.length - 1; i >= 0; i--) y = y * x + c[i]; return y; }
function derivative(c) { var d = []; for (var i = 1; i < c.length; i++) d.push(i * c[i]); return d.length ? d : [0]; }

/** Turning points of polynomial c strictly inside (lo, hi). */
function turningPoints(c, lo, hi) {
  var d = derivative(c), N = 4000, pts = [];
  var step = (hi - lo) / N, x0 = lo, y0 = evalPoly(d, lo);
  for (var i = 1; i <= N; i++) {
    var x1 = lo + i * step, y1 = evalPoly(d, x1);
    if (y0 !== 0 && y1 !== 0 && (y0 < 0) !== (y1 < 0)) {
      var a = x0, b = x1, fa = y0;
      for (var k = 0; k < 80; k++) { var mid = (a + b) / 2, fm = evalPoly(d, mid); if ((fm < 0) === (fa < 0)) { a = mid; fa = fm; } else b = mid; }
      var r = (a + b) / 2;
      if (r > lo + 1e-9 && r < hi - 1e-9) pts.push({ x: r, y: evalPoly(c, r), type: y0 > 0 ? 'max' : 'min' });
    } else if (y1 === 0 && i < N) {
      var y2 = evalPoly(d, x1 + step);
      if ((y0 < 0) !== (y2 < 0)) pts.push({ x: x1, y: evalPoly(c, x1), type: y0 > 0 ? 'max' : 'min' });
      x0 = x1 + step; y0 = y2; i++; continue;
    }
    x0 = x1; y0 = y1;
  }
  return pts;
}

function fmtNum(v, sig) {
  sig = sig || 4;
  if (v === 0) return '0';
  if (!isFinite(v)) return String(v);
  var av = Math.abs(v);
  if (av < 0.001 || av >= 1e7) {
    var e = Math.floor(Math.log(av) / Math.LN10), mant = v / Math.pow(10, e);
    var ms = mant.toPrecision(sig);
    if (Math.abs(parseFloat(ms)) >= 10) { e += 1; ms = (v / Math.pow(10, e)).toPrecision(sig); }
    return ms.replace(/\.?0+$/, '') + '×10^' + e;
  }
  return String(parseFloat(v.toPrecision(Math.max(sig, Math.ceil(Math.log(av + 1) / Math.LN10) + 2))));
}

function fmtShort(v) { return String(Math.round(v * 10) / 10); }

function polyToString(c, sig) {
  var parts = [];
  for (var i = c.length - 1; i >= 0; i--) {
    if (c[i] === 0) continue;
    var num = fmtNum(Math.abs(c[i]), sig || 5);
    var body = i === 0 ? num : ((num === '1' ? '' : num) + (i === 1 ? 'x' : 'x^' + i));
    if (/×/.test(num) && i > 0) body = num + ' ' + (i === 1 ? 'x' : 'x^' + i);
    parts.push((c[i] < 0 ? (parts.length ? ' - ' : '-') : (parts.length ? ' + ' : '')) + body);
  }
  return 'y = ' + (parts.join('') || '0');
}

var SUPERS = ['⁰', '¹', '²', '³', '⁴', '⁵', '⁶', '⁷', '⁸', '⁹'];
function termName(deg) { return deg === 0 ? 'constant' : deg === 1 ? 'x' : 'x' + String(deg).split('').map(function (d) { return SUPERS[+d]; }).join(''); }

/** Checks one submission. sub = raw form answers, st = settings. */
function checkSubmission(sub, st) {
  var checks = {}, feedback = [], key = {};
  CHECKS.forEach(function (c) { checks[c[0]] = { status: 'na', detail: '' }; });
  function set(k, status, detail, hint) { checks[k] = { status: status, detail: detail, hint: hint || detail }; }

  var parsed = null, T = parseNumber(sub.rideEnd);
  try { parsed = parseFactored(sub.factored); }
  catch (e) {
    set('readable', 'fail', 'Could not read the factored form: ' + e.message + '. Fix the typo on the Form Responses tab and Re-check, or check by hand.',
      'We could not read your factored form (' + e.message + '). Check your typing: example y = -0.0000000508(x - 20)(x - 55)^2(x - 100)');
  }

  if (parsed) {
    var c = parsed.coeffs, deg = c.length - 1;
    set('readable', 'pass', 'Read as: ' + polyToString(c));
    key.degree = deg;
    key.aStudent = parsed.a;
    key.zerosText = parsed.zeros.map(function (z) { return fmtShort(z.r) + (z.m > 1 ? ' (mult ' + z.m + ')' : ' (mult 1)'); }).join(', ');
    key.standardText = polyToString(c);

    // a-value
    var px = parseNumber(sub.pointX), py = parseNumber(sub.pointY);
    var prodAt = evalPoly(c, px) / parsed.a; // product of factors at px
    if (isNaN(px) || isNaN(py)) set('aValue', 'warn', 'Point used to find a was not a number.');
    else if (Math.abs(prodAt) < 1e-12) set('aValue', 'fail', 'The point (' + px + ', ' + py + ') is a zero of the polynomial, so it cannot be used to find a.',
      'You cannot use a zero to find a. Use your platform point (0, starting height).');
    else {
      var aExact = py / prodAt;
      key.aExact = aExact;
      var errA = Math.abs(parsed.a - aExact) / Math.abs(aExact);
      if (errA <= st.aTol) set('aValue', 'pass', 'a = ' + fmtNum(parsed.a) + ' (exact ' + fmtNum(aExact) + ', off by ' + (errA * 100).toFixed(2) + '%)');
      else set('aValue', 'fail', 'Student a = ' + fmtNum(parsed.a) + ' but plugging (' + px + ', ' + py + ') in gives a = ' + fmtNum(aExact) +
        ' (' + (errA * 100).toFixed(1) + '% off). ' + (Math.sign(parsed.a) !== Math.sign(aExact) ? 'Sign of a is wrong.' : ''),
        'Your a-value does not make the graph go through (' + px + ', ' + py + '). Re-do the a-value calculation' +
        (Math.sign(parsed.a) !== Math.sign(aExact) ? ' — watch the sign!' : ' — keep more decimal places.'));
    }

    // standard form
    var stu = null;
    try { stu = parseStandard(sub.standard); } catch (e2) {
      set('standard', 'fail', 'Could not read the standard form: ' + e2.message, 'We could not read your standard form (' + e2.message + '). Example format: y = -5.08E-8x^5 + 0.0000183x^4 - 0.00236x^3 + ...');
    }
    if (stu) {
      var bad = [], scale = 0;
      for (var i = 0; i < c.length; i++) scale = Math.max(scale, Math.abs(c[i]) * Math.pow(Math.max(T || 100, 1), i));
      for (i = 0; i < Math.max(c.length, stu.length); i++) {
        var e = c[i] || 0, s = stu[i] || 0;
        var tolAbs = st.coefTol * Math.abs(e) + 1e-9 * scale / Math.pow(Math.max(T || 100, 1), i);
        if (Math.abs(s - e) > tolAbs) bad.push(i);
      }
      if (!bad.length) set('standard', 'pass', 'All ' + (deg + 1) + ' coefficients match within ' + (st.coefTol * 100) + '%.');
      else {
        var names = bad.slice().reverse().map(termName);
        var msg = stu.length - 1 !== deg ? 'Degree is ' + (stu.length - 1) + ' but the factored form has degree ' + deg + '. ' : '';
        set('standard', 'fail', msg + 'These terms do not match: ' + names.join(', ') + '. Correct: ' + polyToString(c),
          msg + 'Your standard form does not match your factored form. Check these terms: ' + names.join(', ') + '.');
      }
    }

    // ride length / domain
    if (isNaN(T)) set('length', 'fail', 'Ride end time is not a number.');
    else if (T >= st.minLength) set('length', 'pass', 'Ride lasts ' + T + ' s (' + (T / 60).toFixed(2) + ' min).');
    else set('length', 'fail', 'Ride lasts only ' + T + ' s; needs ' + st.minLength + ' s.', 'Your ride must last at least ' + st.minLength + ' seconds.');

    // end behavior key
    var lead = c[deg];
    var right = lead > 0 ? '+∞' : '−∞';
    var left = (deg % 2 === 0) ? right : (lead > 0 ? '−∞' : '+∞');
    key.endBehaviorText = 'As x → −∞, y → ' + left + ';  as x → +∞, y → ' + right;
    var sl = String(sub.endLeft), sr = String(sub.endRight);
    var gotL = /\+|rise/.test(sl) ? '+∞' : (/−|-|fall/.test(sl) ? '−∞' : '?');
    var gotR = /\+|rise/.test(sr) ? '+∞' : (/−|-|fall/.test(sr) ? '−∞' : '?');
    if (gotL === left && gotR === right) set('endBeh', 'pass', key.endBehaviorText);
    else set('endBeh', 'fail', 'Student said left ' + gotL + ', right ' + gotR + '. Correct: ' + key.endBehaviorText,
      'Your end behavior is not right. Look at the degree (even/odd) and the sign of the leading coefficient.');

    if (!isNaN(T) && T > 0) {
      var tp = turningPoints(c, 0, T);
      var f0 = evalPoly(c, 0), fT = evalPoly(c, T);
      var ys = [f0, fT].concat(tp.map(function (p) { return p.y; }));
      var minY = Math.min.apply(null, ys), maxY = Math.max.apply(null, ys);
      key.turningText = tp.length ? tp.map(function (p) { return p.type + ' (' + fmtShort(p.x) + ', ' + fmtShort(p.y) + ')'; }).join(';  ') : 'none';
      key.domainText = '[0, ' + T + ']';
      key.rangeText = '[' + fmtShort(minY) + ', ' + fmtShort(maxY) + ']';
      key.endHeightText = fmtShort(fT) + ' ft' + (Math.abs(fT) < 0.5 ? ' (ends on the ground)' : '');
      var bounds = [0].concat(tp.map(function (p) { return p.x; })).concat([T]);
      var inc = [], dec = [];
      for (i = 0; i < bounds.length - 1; i++) {
        var mid = (bounds[i] + bounds[i + 1]) / 2;
        var iv = '(' + fmtShort(bounds[i]) + ', ' + fmtShort(bounds[i + 1]) + ')';
        (evalPoly(derivative(c), mid) > 0 ? inc : dec).push(iv);
      }
      key.increasingText = inc.join(' ∪ ') || 'never';
      key.decreasingText = dec.join(' ∪ ') || 'never';

      // turning points requirement
      if (tp.length >= st.minTurns) set('turns', 'pass', tp.length + ' turning points during the ride.');
      else set('turns', 'fail', 'Only ' + tp.length + ' turning point(s) between 0 and ' + T + ' s; needs ' + st.minTurns + '.',
        'Your coaster needs at least ' + st.minTurns + ' turning points (hills and valleys) during the ride. You have ' + tp.length + '.');

      // tunnel
      if (minY < -0.01) set('tunnel', 'pass', 'Lowest point ' + fmtShort(minY) + ' ft.');
      else set('tunnel', 'fail', 'The ride never goes below ground (lowest point ' + fmtShort(minY) + ' ft).',
        'Your coaster needs a tunnel: part of the ride must go below 0 ft.');

      // platform
      if (f0 > 0) set('platform', 'pass', 'Starts at ' + fmtShort(f0) + ' ft.');
      else set('platform', 'fail', 'Height at time 0 is ' + fmtShort(f0) + ' ft — the ride should start on a platform above ground.',
        'At time 0 your coaster is at ' + fmtShort(f0) + ' ft. It should start on a platform above the ground.');

      // realistic
      var probs = [];
      if (maxY > st.maxHeight) probs.push('peak ' + fmtShort(maxY) + ' ft is taller than ' + st.maxHeight + ' ft');
      if (minY < st.minDepth) probs.push('tunnel ' + fmtShort(minY) + ' ft is deeper than ' + st.minDepth + ' ft');
      if (probs.length) set('realistic', 'warn', 'Unrealistic: ' + probs.join('; ') + '.', 'Is your coaster realistic? ' + probs.join('; ') + '.');
      else set('realistic', 'pass', 'Heights from ' + fmtShort(minY) + ' to ' + fmtShort(maxY) + ' ft.');

      // turning point answer
      var given = String(sub.turning === undefined ? '' : sub.turning).split(/[,;\s]+|and/).map(parseNumber).filter(function (v) { return !isNaN(v); }).sort(function (u, v) { return u - v; });
      var tolT = Math.max(st.timeTol, 0.01 * T);
      var okT = given.length === tp.length && given.every(function (g, j) { return Math.abs(g - tp[j].x) <= tolT; });
      var keyTimes = tp.map(function (p) { return fmtShort(p.x); }).join(', ');
      if (okT) set('turnAns', 'pass', 'Student: ' + given.join(', ') + ' (key: ' + keyTimes + ')');
      else set('turnAns', 'fail', 'Student: ' + (given.join(', ') || '(none)') + '  Key: ' + keyTimes,
        given.length !== tp.length ? 'You listed ' + given.length + ' turning point time(s), but your polynomial has ' + tp.length + ' during the ride.' :
          'At least one of your turning point times does not match your polynomial. Use Desmos to click on each max and min.');

      // range answer
      var lo = parseNumber(sub.rangeLow), hi = parseNumber(sub.rangeHigh);
      var tolH = Math.max(st.heightTol, 0.02 * (maxY - minY));
      if (!isNaN(lo) && !isNaN(hi) && Math.abs(lo - minY) <= tolH && Math.abs(hi - maxY) <= tolH)
        set('rangeAns', 'pass', 'Student: [' + lo + ', ' + hi + '] (key ' + key.rangeText + ')');
      else set('rangeAns', 'fail', 'Student: [' + sub.rangeLow + ', ' + sub.rangeHigh + ']  Key: ' + key.rangeText,
        'Your range does not match. Remember the range is only for the ride (from time 0 to ' + T + ' s) — check the endpoints too, not just the turning points.');
    }

    // repeated zero
    if (st.requireRepeated) {
      var rep = parsed.zeros.filter(function (z) { return z.m >= 2; });
      if (rep.length) set('repeated', 'pass', 'Repeated zero at ' + rep.map(function (z) { return fmtShort(z.r); }).join(', '));
      else set('repeated', 'fail', 'No zero has multiplicity 2 or more.', 'Your coaster needs at least one zero with multiplicity 2 or more (the track touches the ground and bounces back).');
    }
  }

  var applicable = 0, pass = 0, fails = [];
  CHECKS.forEach(function (ck) {
    var s = checks[ck[0]].status;
    if (s === 'na') return;
    applicable++;
    if (s === 'pass') pass++; else fails.push(ck[1]);
    if (s !== 'pass') feedback.push(ck[1] + ': ' + checks[ck[0]].detail);
  });
  if (!parsed) { applicable = CHECKS.length - (st.requireRepeated ? 0 : 1); }
  var summary = !parsed ? 'Could not read equation — check by hand' :
    (fails.length ? 'Fix: ' + fails.join(', ') : 'All requirements met 🎢');
  return { checks: checks, key: key, passCount: pass, applicable: applicable, summary: summary, feedback: feedback };
}

// Allows the math core to be tested outside Google (ignored by Apps Script).
if (typeof module !== 'undefined') module.exports = { parseFactored: parseFactored, parseStandard: parseStandard, polyToString: polyToString, checkSubmission: checkSubmission, turningPoints: turningPoints, evalPoly: evalPoly, fmtNum: fmtNum, normalizeExpr: normalizeExpr, parseNumber: parseNumber };
