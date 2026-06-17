/* ==========================================================================
 *  Dynasty Group Delete Solutions - v4.0
 *  MyGeotab Add-In (vanilla JS, CSP / Trusted Types compliant)
 *
 *  NEW in v4.0 (UX restored from the original app):
 *   - On group select, the add-in SCANS (read-only) and lists the group's
 *     related entities by category (devices, users, drivers, rules, zones,
 *     exception rules, marketplace add-ins). Each category has a checkbox.
 *   - The Delete button stays disabled until every listed category is checked
 *     AND the "irreversible" confirm is checked. If a group has no references,
 *     it can be deleted directly after the confirm.
 *   - The preview is purely read-only (Get calls) so selecting a group never
 *     deletes anything. The actual Delete still runs the robust loop below
 *     (Remove -> parse GroupRelationViolatedException -> auto-clear -> repeat),
 *     so anything the preview can't see is still handled at delete time.
 *
 *  Front-end: self-contained Dynasty Safety Console "NOC" theme (dark + light
 *  toggle), styling injected via injectStyles() - no external styles.css.
 *
 *  CSP rules: createElement / createElementNS / textContent only (no
 *  innerHTML), classList + addEventListener (no onclick attributes), var only
 *  (IIFE scope), no literal newlines inside JS strings.
 * ========================================================================== */
geotab.addin.dynastyGroupDelete = function () {
  'use strict';

  var api = null;
  var creds = null;
  var elRoot = null;
  var elApp = null;
  var elThemeBtn = null;
  var elSelect = null;
  var elRefsTitle = null;
  var elRefs = null;
  var elConfirm = null;
  var elConfirmBox = null;
  var elDeleteBtn = null;
  var elRefreshBtn = null;
  var elLog = null;
  var busy = false;
  var scanned = false;
  var catChecks = [];

  var MAX_NAMES = 6;
  var SYSTEM_PREFIX = 'Group';
  var COMPANY_ID = 'GroupCompanyId';
  var fallbackGroupId = COMPANY_ID;

  var CAT_MAP = {
    devices: ['Device', ['groups']],
    users: ['User', ['companyGroups']],
    drivers: ['User', ['driverGroups']],
    rules: ['Rule', ['groups']],
    zones: ['Zone', ['groups']],
    eventRules: ['ExceptionRule', ['groups']],
    dvirLogs: ['DVIRLog', ['groups']],
    defects: ['DefectRemark', ['groups']],
    inspectionTemplates: ['DVIRDefect', ['groups']],
    maintenanceWorkJobs: ['MaintenanceEvent', ['groups']],
    customReportSchedules: ['CustomReportSchedule', ['scopeGroups', 'groups']]
  };

  // Fields Geotab requires to be non-empty. If removing the group would leave
  // one empty, fall back to the Company group so the Set is accepted.
  var NON_EMPTY_FIELDS = { companyGroups: true };

  // ---------------------------- styling -----------------------------------
  function injectStyles() {
    if (document.getElementById('cdg-styles')) { return; }
    var css = [
      '.cdg-app{',
      '  --bg:#0d1117; --panel:#161b22; --panel-2:#1c232c; --border:#2a323d;',
      '  --text:#e6edf3; --muted:#8b98a5; --amber:#ffb454; --amber-strong:#f59e1b;',
      '  --red:#ff5d5d; --red-strong:#e3534f; --blue:#5ab0ff; --green:#46af55;',
      '  --log-bg:#080b10; --radius:10px;',
      '  --font:"Segoe UI",Roboto,system-ui,-apple-system,"Helvetica Neue",Arial,sans-serif;',
      '  --mono:"Roboto Mono",ui-monospace,"SFMono-Regular",Menlo,Consolas,monospace;',
      '  box-sizing:border-box; min-height:100vh; background:var(--bg);',
      '  color:var(--text); font-family:var(--font); padding:20px 24px 40px;',
      '}',
      '.cdg-app *{box-sizing:border-box}',
      '.cdg-app.cdg-light{',
      '  --bg:#f4f6f8; --panel:#ffffff; --panel-2:#eef2f6; --border:#d6dee6;',
      '  --text:#1f2833; --muted:#4e677e; --log-bg:#0d1117;',
      '}',
      '.cdg-header{display:flex;align-items:center;gap:14px;margin-bottom:18px}',
      '.cdg-shield{width:34px;height:34px;flex:0 0 auto}',
      '.cdg-brandtext{display:flex;flex-direction:column;gap:2px;flex:1 1 auto;min-width:0}',
      '.cdg-title{font-size:20px;font-weight:600;line-height:24px;margin:0;color:var(--text);letter-spacing:.2px}',
      '.cdg-eyebrow{font-size:11px;font-weight:600;letter-spacing:1.4px;text-transform:uppercase;color:var(--amber);margin:0}',
      '.cdg-theme{flex:0 0 auto;border:1px solid var(--border);background:var(--panel-2);color:var(--muted);font-family:var(--font);font-size:12px;font-weight:500;padding:6px 12px;border-radius:999px;cursor:pointer;outline:none}',
      '.cdg-theme:hover{color:var(--text);border-color:var(--amber)}',
      '.cdg-card{background:var(--panel);border:1px solid var(--border);border-radius:var(--radius);padding:20px;max-width:760px;box-shadow:0 1px 0 rgba(0,0,0,.25)}',
      '.cdg-intro{font-size:13px;line-height:20px;color:var(--muted);margin:0 0 16px}',
      '.cdg-row{display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap}',
      '.cdg-field{flex:1 1 320px;min-width:240px}',
      '.cdg-label{display:block;font-size:11px;font-weight:600;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin-bottom:6px}',
      '.cdg-select{width:100%;height:38px;padding:8px 12px;font-family:var(--font);font-size:13px;color:var(--text);background:var(--panel-2);border:1px solid var(--border);border-radius:8px;outline:none;cursor:pointer;-webkit-appearance:none;-moz-appearance:none;appearance:none}',
      '.cdg-select:hover{border-color:var(--amber)}',
      '.cdg-select:focus{border-color:var(--amber);box-shadow:0 0 0 2px rgba(255,180,84,.18)}',
      '.cdg-actions{display:flex;gap:8px;flex-wrap:wrap}',
      '.cdg-btn{display:inline-flex;align-items:center;justify-content:center;height:38px;padding:0 16px;font-family:var(--font);font-size:13px;font-weight:600;color:var(--text);background:var(--panel-2);border:1px solid var(--border);border-radius:8px;cursor:pointer;outline:none}',
      '.cdg-btn:hover{border-color:var(--amber);color:var(--amber)}',
      '.cdg-btn:disabled,.cdg-btn[disabled]{opacity:.45;cursor:default;pointer-events:none}',
      '.cdg-btn-danger{color:#fff;background:var(--red);border-color:var(--red)}',
      '.cdg-btn-danger:hover{color:#fff;background:var(--red-strong);border-color:var(--red-strong)}',
      '.cdg-btn-danger:not([disabled]){box-shadow:0 0 0 2px rgba(255,93,93,.25)}',
      // references preview
      '.cdg-refs-title{font-size:11px;font-weight:600;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin:18px 0 8px;display:none}',
      '.cdg-refs{display:flex;flex-direction:column;gap:8px}',
      '.cdg-scan{font-size:12px;color:var(--muted);font-style:italic;padding:4px 0}',
      '.cdg-empty{font-size:13px;color:var(--green);background:rgba(70,175,85,.1);border:1px solid rgba(70,175,85,.4);border-radius:8px;padding:10px 12px}',
      '.cdg-ref-cat{display:flex;gap:10px;align-items:flex-start;padding:10px 12px;background:var(--panel-2);border:1px solid var(--border);border-radius:8px}',
      '.cdg-ref-cat.warn{background:rgba(255,180,84,.08);border-color:rgba(255,180,84,.45)}',
      '.cdg-ref-body{flex:1 1 auto;min-width:0}',
      '.cdg-ref-head{display:flex;align-items:center;gap:8px;margin-bottom:3px}',
      '.cdg-ref-label{font-size:13px;font-weight:600;color:var(--text)}',
      '.cdg-ref-count{font-size:11px;font-weight:600;color:var(--amber);background:rgba(255,180,84,.14);border-radius:999px;padding:1px 8px}',
      '.cdg-ref-note{font-size:11px;color:var(--amber);margin-left:2px}',
      '.cdg-ref-names{font-size:12px;color:var(--muted);line-height:1.5;word-break:break-word}',
      // checkbox
      '.cdg-check{position:relative;flex:0 0 auto;width:18px;height:18px;margin-top:1px}',
      '.cdg-check input{position:absolute;opacity:0;width:18px;height:18px;margin:0;cursor:pointer}',
      '.cdg-box{width:18px;height:18px;border:1px solid var(--border);border-radius:4px;background:var(--panel);display:flex;align-items:center;justify-content:center}',
      '.cdg-check input:checked + .cdg-box{background:var(--amber);border-color:var(--amber)}',
      '.cdg-box svg{display:none;width:12px;height:12px}',
      '.cdg-check input:checked + .cdg-box svg{display:block}',
      '.cdg-check input:focus + .cdg-box{box-shadow:0 0 0 2px rgba(255,180,84,.3)}',
      // confirm
      '.cdg-confirm{display:none;gap:10px;align-items:flex-start;margin:16px 0 4px;padding:12px 14px;background:rgba(255,93,93,.08);border:1px solid rgba(255,93,93,.35);border-radius:8px}',
      '.cdg-confirm.show{display:flex}',
      '.cdg-confirm .cdg-check input:checked + .cdg-box{background:var(--red);border-color:var(--red)}',
      '.cdg-confirm-text{font-size:12px;line-height:18px;color:var(--text);cursor:pointer}',
      '.cdg-confirm-text b{color:var(--red)}',
      // log
      '.cdg-logtitle{font-size:11px;font-weight:600;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin:18px 0 8px}',
      '.cdg-log{background:var(--log-bg);border:1px solid var(--border);color:#cdd6e0;font-family:var(--mono);font-size:12px;line-height:1.7;padding:12px 14px;border-radius:8px;max-height:300px;overflow-y:auto}',
      '.cdg-log-empty{color:#5b6673;font-style:italic}',
      '.cdg-log-line{white-space:pre-wrap;word-break:break-word}',
      '.cdg-log-ok{color:var(--green)}',
      '.cdg-log-err{color:var(--red)}',
      '.cdg-log-info{color:#8fb9e8}',
      '.cdg-disclaimer{margin-top:16px;font-size:12px;line-height:1.5;color:var(--muted);background:rgba(143,185,232,.08);border:1px solid rgba(143,185,232,.3);border-radius:8px;padding:10px 12px}',
      '.cdg-disclaimer a{color:var(--blue);text-decoration:none}',
      '.cdg-disclaimer a:hover{text-decoration:underline}',
      '.cdg-footer{max-width:760px;margin-top:14px;font-size:11px;color:var(--muted);text-align:right;letter-spacing:.3px}'
    ].join('\n');

    var style = document.createElement('style');
    style.id = 'cdg-styles';
    style.setAttribute('type', 'text/css');
    style.appendChild(document.createTextNode(css));
    (document.head || document.getElementsByTagName('head')[0] || document.documentElement).appendChild(style);
  }

  // ----------------------------- helpers ---------------------------------
  function endpoint() {
    return window.location.origin + '/apiv1';
  }

  function rpc(method, params) {
    var body = { method: method, params: params || {} };
    if (creds) { body.params.credentials = creds; }
    return fetch(endpoint(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }).then(function (res) {
      return res.json();
    }).then(function (j) {
      if (j && j.error) {
        var inner = (j.error.errors && j.error.errors[0]) || {};
        var err = new Error(inner.message || j.error.message || 'API error');
        err.geotabName = inner.name || (j.error.data && j.error.data.type) || '';
        err.dataInfo = (typeof inner.data === 'string') ? inner.data :
          (j.error.data && j.error.data.info) || null;
        throw err;
      }
      return j.result;
    });
  }

  function getEntities(typeName, search) {
    var p = { typeName: typeName };
    if (search) { p.search = search; }
    return rpc('Get', p);
  }

  function safeGet(typeName, gid) {
    return getEntities(typeName, { groups: [{ id: gid }] }).then(function (rows) {
      return rows || [];
    }).catch(function () { return []; });
  }

  function hasGroup(arr, gid) {
    var i;
    if (!arr) { return false; }
    for (i = 0; i < arr.length; i++) {
      if (arr[i] && arr[i].id === gid) { return true; }
    }
    return false;
  }

  function isSystemGroup(id) {
    return id === COMPANY_ID || String(id).indexOf(SYSTEM_PREFIX) === 0;
  }

  // Keep only entities whose own groups[] DIRECTLY contains gid (i.e. they
  // actually reference this group), not everything merely visible in scope.
  function refsOf(rows, gid) {
    var out = [];
    (rows || []).forEach(function (e) {
      if (e && hasGroup(e.groups, gid)) { out.push(e); }
    });
    return out;
  }

  // Find the parent group of gid (the group whose children include gid).
  // Used as the fallback scope for users left with no group, so a user scoped
  // only to the deleted group moves UP one level - not to organization-wide.
  function resolveParent(gid) {
    return getEntities('Group').then(function (groups) {
      var i, j, g;
      for (i = 0; i < groups.length; i++) {
        g = groups[i];
        if (g && g.children) {
          for (j = 0; j < g.children.length; j++) {
            if (g.children[j] && g.children[j].id === gid) { return g.id; }
          }
        }
      }
      return COMPANY_ID;
    }).catch(function () { return COMPANY_ID; });
  }

  function entName(e) {
    if (!e) { return ''; }
    return e.name || e.serialNumber || (e.firstName ? (e.firstName + ' ' + (e.lastName || '')).trim() : '') || e.id || '';
  }

  function namesLine(names) {
    var shown = names.slice(0, MAX_NAMES).join(', ');
    if (names.length > MAX_NAMES) { shown = shown + '  +' + (names.length - MAX_NAMES) + ' more'; }
    return shown;
  }

  // ------------------------------- UI -------------------------------------
  var SVG_NS = 'http://www.w3.org/2000/svg';

  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) { node.className = className; }
    if (text != null) { node.textContent = text; }
    return node;
  }

  function svgEl(name, attrs) {
    var node = document.createElementNS(SVG_NS, name);
    var k;
    for (k in attrs) {
      if (attrs.hasOwnProperty(k)) { node.setAttribute(k, attrs[k]); }
    }
    return node;
  }

  function shieldIcon() {
    var svg = svgEl('svg', { 'class': 'cdg-shield', viewBox: '0 0 32 32' });
    svg.appendChild(svgEl('path', {
      d: 'M16 2 L28 6 V15 C28 22.5 22.8 27.6 16 30 C9.2 27.6 4 22.5 4 15 V6 Z',
      fill: 'var(--amber)'
    }));
    svg.appendChild(svgEl('path', {
      d: 'M11 16 l3.4 3.4 L21 12.8',
      fill: 'none', stroke: '#0d1117', 'stroke-width': '2.6',
      'stroke-linecap': 'round', 'stroke-linejoin': 'round'
    }));
    return svg;
  }

  function checkMark() {
    var svg = svgEl('svg', { viewBox: '0 0 16 16' });
    svg.appendChild(svgEl('path', {
      d: 'M6.5 11.5 L3 8 l1.4 -1.4 L6.5 8.7 l5.1 -5.1 L13 4 z',
      fill: '#ffffff'
    }));
    return svg;
  }

  // Builds a label-wrapped checkbox; returns the <input>.
  function makeCheck(onChange) {
    var wrap = el('label', 'cdg-check');
    var input = el('input');
    input.type = 'checkbox';
    input.addEventListener('change', onChange);
    var box = el('span', 'cdg-box');
    box.appendChild(checkMark());
    wrap.appendChild(input);
    wrap.appendChild(box);
    input._wrap = wrap;
    return input;
  }

  function toggleTheme() {
    if (!elApp) { return; }
    if (elApp.classList.contains('cdg-light')) {
      elApp.classList.remove('cdg-light');
      elThemeBtn.textContent = 'Light';
    } else {
      elApp.classList.add('cdg-light');
      elThemeBtn.textContent = 'Dark';
    }
  }

  function setBusy(state) {
    busy = state;
    if (elSelect) { elSelect.disabled = state; }
    if (elRefreshBtn) { elRefreshBtn.disabled = state; }
    syncDeleteEnabled();
  }

  function allCatsChecked() {
    var i;
    for (i = 0; i < catChecks.length; i++) {
      if (!catChecks[i].checked) { return false; }
    }
    return true;
  }

  function syncDeleteEnabled() {
    if (!elDeleteBtn) { return; }
    var ok = !busy && scanned && elSelect && elSelect.value &&
      elConfirm && elConfirm.checked && allCatsChecked();
    elDeleteBtn.disabled = !ok;
  }

  function log(msg, kind) {
    if (!elLog) { return; }
    var empty = elLog.querySelector('.cdg-log-empty');
    if (empty) { elLog.removeChild(empty); }
    var line = el('div', 'cdg-log-line' + (kind ? ' cdg-log-' + kind : ''));
    line.textContent = '[' + new Date().toLocaleTimeString() + ']  ' + msg;
    elLog.appendChild(line);
    elLog.scrollTop = elLog.scrollHeight;
  }

  function clearLog() {
    if (!elLog) { return; }
    while (elLog.firstChild) { elLog.removeChild(elLog.firstChild); }
  }

  function logPlaceholder() {
    if (!elLog) { return; }
    clearLog();
    elLog.appendChild(el('div', 'cdg-log-empty', 'Waiting for action...'));
  }

  function clearNode(node) {
    while (node && node.firstChild) { node.removeChild(node.firstChild); }
  }

  function buildUI() {
    while (elRoot.firstChild) { elRoot.removeChild(elRoot.firstChild); }

    elApp = el('div', 'cdg-app');

    var header = el('div', 'cdg-header');
    header.appendChild(shieldIcon());
    var brandText = el('div', 'cdg-brandtext');
    brandText.appendChild(el('div', 'cdg-eyebrow', 'Dynasty Communications'));
    brandText.appendChild(el('h1', 'cdg-title', 'Dynasty Group Deletion'));
    header.appendChild(brandText);
    elThemeBtn = el('button', 'cdg-theme', 'Light');
    elThemeBtn.type = 'button';
    elThemeBtn.addEventListener('click', toggleTheme);
    header.appendChild(elThemeBtn);
    elApp.appendChild(header);

    var card = el('div', 'cdg-card');
    card.appendChild(el('p', 'cdg-intro', 'Select a group to see everything that references it. Tick each category to acknowledge it, then confirm to delete. References are cleared automatically; marketplace add-ins that block deletion are reported for your provider to disable in MyAdmin.'));

    var row = el('div', 'cdg-row');
    var field = el('div', 'cdg-field');
    field.appendChild(el('label', 'cdg-label', 'Group'));
    elSelect = el('select', 'cdg-select');
    var ph = el('option', null, 'Loading groups...');
    ph.value = '';
    elSelect.appendChild(ph);
    elSelect.addEventListener('change', onSelectChange);
    field.appendChild(elSelect);
    row.appendChild(field);

    var actions = el('div', 'cdg-actions');
    elRefreshBtn = el('button', 'cdg-btn', 'Refresh');
    elRefreshBtn.type = 'button';
    elRefreshBtn.addEventListener('click', loadGroups);
    actions.appendChild(elRefreshBtn);
    row.appendChild(actions);
    card.appendChild(row);

    // references preview
    elRefsTitle = el('div', 'cdg-refs-title', 'Related entities');
    card.appendChild(elRefsTitle);
    elRefs = el('div', 'cdg-refs');
    card.appendChild(elRefs);

    // confirm
    var confirm = el('div', 'cdg-confirm');
    elConfirm = makeCheck(syncDeleteEnabled);
    elConfirm.id = 'cdg-confirm-input';
    var confirmText = el('label', 'cdg-confirm-text');
    confirmText.setAttribute('for', 'cdg-confirm-input');
    confirmText.appendChild(el('b', null, 'Cascade delete is irreversible.'));
    confirmText.appendChild(document.createTextNode(' This permanently deletes the selected group and removes it from all referencing entities.'));
    confirm.appendChild(elConfirm._wrap);
    confirm.appendChild(confirmText);
    elConfirmBox = confirm;
    card.appendChild(confirm);

    // delete button (below the checks, like the original app)
    var delRow = el('div', 'cdg-actions');
    elDeleteBtn = el('button', 'cdg-btn cdg-btn-danger', 'Delete group');
    elDeleteBtn.type = 'button';
    elDeleteBtn.disabled = true;
    elDeleteBtn.addEventListener('click', onDeleteClick);
    delRow.appendChild(elDeleteBtn);
    card.appendChild(delRow);

    card.appendChild(el('div', 'cdg-logtitle', 'Execution log'));
    elLog = el('div', 'cdg-log');
    card.appendChild(elLog);

    var disc = el('div', 'cdg-disclaimer');
    disc.appendChild(document.createTextNode('Note: If the groups are not deleting, please contact Dynasty support at '));
    var mail = el('a', null, 'support@dynastync.com');
    mail.setAttribute('href', 'mailto:support@dynastync.com');
    disc.appendChild(mail);
    disc.appendChild(document.createTextNode('.'));
    card.appendChild(disc);

    elApp.appendChild(card);
    elApp.appendChild(el('div', 'cdg-footer', 'Dynasty Communications'));

    elRoot.appendChild(elApp);
    logPlaceholder();
  }

  function resetPreview() {
    scanned = false;
    catChecks = [];
    clearNode(elRefs);
    if (elConfirm) { elConfirm.checked = false; }
    if (elRefsTitle) { elRefsTitle.style.display = 'none'; }
    if (elConfirmBox) { elConfirmBox.classList.remove('show'); }
    syncDeleteEnabled();
  }

  function onSelectChange() {
    resetPreview();
    if (elSelect.value) { scanGroup(elSelect.value); }
  }

  function loadGroups() {
    if (busy) { return; }
    resetPreview();
    log('Loading groups...', 'info');
    getEntities('Group').then(function (groups) {
      while (elSelect.firstChild) { elSelect.removeChild(elSelect.firstChild); }
      var ph = el('option', null, '-- Select a group --');
      ph.value = '';
      elSelect.appendChild(ph);

      groups.filter(function (g) {
        return !isSystemGroup(g.id);
      }).sort(function (a, b) {
        return String(a.name || '').localeCompare(String(b.name || ''));
      }).forEach(function (g) {
        var opt = el('option', null, (g.name || g.id));
        opt.value = g.id;
        elSelect.appendChild(opt);
      });
      log('Ready. ' + (elSelect.options.length - 1) + ' deletable groups loaded.', 'ok');
      syncDeleteEnabled();
    }).catch(function (e) {
      log('Could not load groups: ' + e.message, 'err');
    });
  }

  // ----------------------- read-only reference scan -----------------------
  function scanGroup(gid) {
    elRefsTitle.style.display = 'block';
    clearNode(elRefs);
    elRefs.appendChild(el('div', 'cdg-scan', 'Scanning references for this group...'));

    var pDevices = safeGet('Device', gid);
    var pRules = safeGet('Rule', gid);
    var pZones = safeGet('Zone', gid);
    var pExc = safeGet('ExceptionRule', gid);
    var pUsers = getEntities('User').catch(function () { return []; });
    var pAddins = getEntities('AddIn').catch(function () { return []; });

    Promise.all([pDevices, pRules, pZones, pExc, pUsers, pAddins]).then(function (res) {
      // guard: selection may have changed while scanning
      if (elSelect.value !== gid) { return; }

      var cats = [];
      pushCat(cats, 'Devices', refsOf(res[0], gid));
      pushCat(cats, 'Rules', refsOf(res[1], gid));
      pushCat(cats, 'Zones', refsOf(res[2], gid));
      pushCat(cats, 'Exception rules', refsOf(res[3], gid));

      var members = [];
      var drivers = [];
      var advanced = [];
      (res[4] || []).forEach(function (u) {
        var inCompany = hasGroup(u.companyGroups, gid);
        if (inCompany && filterRefsGroup(u.accessGroupFilter, gid)) {
          advanced.push(entName(u));
        } else if (inCompany) {
          members.push(entName(u));
        }
        if (hasGroup(u.driverGroups, gid)) { drivers.push(entName(u)); }
      });
      if (members.length) { cats.push({ label: 'Users (data access)', names: members }); }
      if (drivers.length) { cats.push({ label: 'Drivers', names: drivers }); }
      if (advanced.length) { cats.push({ label: 'Users assigned only to this group', names: advanced, warn: true, note: 'assign these users another group first' }); }

      var addins = [];
      (res[5] || []).forEach(function (a) {
        if (hasGroup(a.groups, gid)) {
          var nm = (a.configuration && a.configuration.name) || a.name || a.id;
          addins.push(nm + '  (auto-enroll: ' + (a.isAutoEnrollEnabled ? 'ON' : 'OFF') + ')');
        }
      });
      if (addins.length) { cats.push({ label: 'Marketplace add-ins', names: addins, warn: true, note: 'cleared in MyAdmin if it blocks delete' }); }

      renderRefs(cats);
    });
  }

  function pushCat(cats, label, rows) {
    if (rows && rows.length) {
      cats.push({ label: label, names: rows.map(entName) });
    }
  }

  function renderRefs(cats) {
    clearNode(elRefs);
    catChecks = [];

    if (!cats.length) {
      elRefs.appendChild(el('div', 'cdg-empty', 'No related entities found. This group can be deleted directly.'));
    } else {
      cats.forEach(function (c) {
        var cat = el('div', 'cdg-ref-cat' + (c.warn ? ' warn' : ''));
        var chk = makeCheck(syncDeleteEnabled);
        catChecks.push(chk);
        cat.appendChild(chk._wrap);

        var body = el('div', 'cdg-ref-body');
        var head = el('div', 'cdg-ref-head');
        head.appendChild(el('span', 'cdg-ref-label', c.label));
        head.appendChild(el('span', 'cdg-ref-count', String(c.names.length)));
        if (c.note) { head.appendChild(el('span', 'cdg-ref-note', c.note)); }
        body.appendChild(head);
        body.appendChild(el('div', 'cdg-ref-names', namesLine(c.names)));
        cat.appendChild(body);
        elRefs.appendChild(cat);
      });
    }

    scanned = true;
    elConfirmBox.classList.add('show');
    syncDeleteEnabled();
  }

  function onDeleteClick() {
    var gid = elSelect.value;
    if (!gid) { log('Please select a group first.', 'err'); return; }
    if (!elConfirm.checked || !allCatsChecked()) {
      log('Please review and tick all boxes before deleting.', 'err');
      return;
    }
    var name = elSelect.options[elSelect.selectedIndex].textContent;
    clearLog();
    setBusy(true);
    log('Starting delete for "' + name + '" (' + gid + ')', 'info');
    resolveParent(gid).then(function (parentId) {
      fallbackGroupId = parentId;
      if (parentId === COMPANY_ID) {
        log('Note: this is a top-level group. Any user scoped only to it will be moved to the Company Group (organization-wide access) - reassign them afterward if narrower access is needed.', 'info');
      } else {
        log('Any user scoped only to this group will be moved up to its parent group (not full access).', 'info');
      }
      runDelete(gid, name, 1);
    });
  }

  // ----------------------- core delete (loop) -----------------------------
  function runDelete(gid, name, iter) {
    var MAX_ITERS = 8;
    if (iter > MAX_ITERS) {
      log('Stopped: too many iterations. Remaining blockers need manual handling.', 'err');
      setBusy(false);
      return;
    }

    rpc('Remove', { typeName: 'Group', entity: { id: gid } }).then(function () {
      log('"' + name + '" was deleted successfully.', 'ok');
      resetPreview();
      setBusy(false);
      loadGroups();
    }).catch(function (err) {
      if (err.geotabName !== 'GroupRelationViolatedException' || !err.dataInfo) {
        log('Delete failed: ' + err.message, 'err');
        setBusy(false);
        return;
      }
      var info;
      try { info = JSON.parse(err.dataInfo); } catch (e) { info = null; }
      if (!info) {
        log('Could not parse relations. ' + err.message, 'err');
        setBusy(false);
        return;
      }
      handleBlockers(gid, name, info, iter);
    });
  }

  function handleBlockers(gid, name, info, iter) {
    var tasks = [];
    var addInBlockers = [];
    var manual = [];
    var entMap = {};
    var k, cat, mapKey;

    for (k in info) {
      if (!info.hasOwnProperty(k) || k === 'group') { continue; }
      cat = info[k];
      if (!Array.isArray(cat) || cat.length === 0) { continue; }

      if (k === 'addIns') {
        addInBlockers = cat.slice();
      } else if (k === 'groupFilters') {
        tasks.push(clearGroupFilters(cat, gid));
      } else if (CAT_MAP[k]) {
        (function (typeName, fields) {
          cat.forEach(function (item) {
            var key = typeName + '|' + item.id;
            if (!entMap[key]) { entMap[key] = { typeName: typeName, id: item.id, fields: [] }; }
            fields.forEach(function (f) {
              if (entMap[key].fields.indexOf(f) < 0) { entMap[key].fields.push(f); }
            });
          });
        })(CAT_MAP[k][0], CAT_MAP[k][1]);
      } else {
        manual.push(k + ' (' + cat.length + ')');
      }
    }

    for (mapKey in entMap) {
      if (entMap.hasOwnProperty(mapKey)) {
        tasks.push(clearEntity(entMap[mapKey].typeName, entMap[mapKey].fields, entMap[mapKey].id, gid));
      }
    }

    if (tasks.length === 0 && addInBlockers.length === 0 && manual.length === 0) {
      log('No clearable references found but delete still blocked.', 'err');
      setBusy(false);
      return;
    }

    if (tasks.length === 0) {
      reportUnclearable(addInBlockers, manual, name);
      setBusy(false);
      return;
    }

    log('Iteration ' + iter + ': clearing ' + tasks.length + ' linked item(s)...', 'info');
    var wrapped = tasks.map(function (t) {
      return t.then(function (r) {
        return r === false ? false : (r ? true : null);
      }).catch(function (e) {
        log('  ' + (e && e.message ? e.message : 'clear failed'), 'err');
        return false;
      });
    });
    Promise.all(wrapped).then(function (results) {
      if (results.indexOf(true) >= 0) {
        runDelete(gid, name, iter + 1);
      } else {
        log('Could not finish. The user(s) listed above only have this group assigned - please assign them another group in MyGeotab, then run Delete again. If it continues, contact Dynasty support.', 'err');
        setBusy(false);
      }
    });
  }

  function setEntity(typeName, ent, label) {
    return rpc('Set', { typeName: typeName, entity: ent }).then(function () {
      return true;
    }).catch(function (e) {
      var msg = e.message || 'Set failed';
      if (typeName === 'User' && /AccessGroupFilter|CompanyGroups/i.test(msg)) {
        log('  User "' + label + '" only has this group assigned. Please assign this user another group in MyGeotab, then run Delete again.', 'err');
      } else {
        log('  ' + typeName + ' "' + label + '": ' + msg, 'err');
      }
      return false;
    });
  }

  function filterRefsGroup(f, gid) {
    if (!f || !f.groupFilterConditions) { return false; }
    var i, c;
    for (i = 0; i < f.groupFilterConditions.length; i++) {
      c = f.groupFilterConditions[i];
      if (c) {
        if (c.groupId === gid) { return true; }
        if (c.groupFilterConditions && filterRefsGroup(c, gid)) { return true; }
      }
    }
    return false;
  }

  function clearEntity(typeName, fields, id, gid) {
    return getEntities(typeName, { id: id }).then(function (rows) {
      var ent = rows && rows[0];
      if (!ent) { return null; }
      var label = ent.name || id;

      // A User whose advanced Data Access (accessGroupFilter) is locked to this
      // group cannot be re-scoped via the API - Geotab does not allow writing
      // accessGroupFilter. Flag it for a one-time manual fix and skip the
      // (doomed) Set, so the log stays accurate.
      if (typeName === 'User' && filterRefsGroup(ent.accessGroupFilter, gid)) {
        log('  User "' + label + '" only has this group assigned. Please assign this user another group in MyGeotab, then run Delete again.', 'err');
        return false;
      }

      var changed = false;
      fields.forEach(function (f) {
        if (Array.isArray(ent[f]) && hasGroup(ent[f], gid)) {
          var filtered = ent[f].filter(function (g) { return g.id !== gid; });
          if (filtered.length === 0 && NON_EMPTY_FIELDS[f]) {
            filtered = [{ id: fallbackGroupId }];
            log('  ' + typeName + ' "' + label + '": ' + f + ' would be empty - moved to parent group', 'info');
          }
          ent[f] = filtered;
          changed = true;
        }
      });
      if (!changed) { return null; }
      return setEntity(typeName, ent, label).then(function (ok) {
        if (ok) { log('  cleared ' + typeName + ' "' + label + '"', 'info'); }
        return ok;
      });
    });
  }

  function clearGroupFilters(items, gid) {
    return getEntities('CustomReportSchedule').then(function (crs) {
      var chain = Promise.resolve();
      items.forEach(function (f) {
        chain = chain.then(function () {
          return getEntities('GroupFilter', { id: f.id }).then(function (rows) {
            var gf = rows && rows[0];
            if (!gf || !gf.groupFilterCondition) { return null; }
            var conds = (gf.groupFilterCondition.groupFilterConditions || [])
              .filter(function (c) { return c.groupId !== gid; });
            if (conds.length === 1) {
              gf.groupFilterCondition = { isNegated: false, groupId: conds[0].groupId };
            } else if (conds.length >= 2) {
              gf.groupFilterCondition.groupFilterConditions = conds;
            } else {
              gf.groupFilterCondition = null;
            }
            log('  cleared report group-filter ' + f.id, 'info');
            return rpc('Set', { typeName: 'GroupFilter', entity: gf }).catch(function () {
              var rep = (crs || []).filter(function (r) {
                return r.scopeGroupFilter && r.scopeGroupFilter.id === f.id;
              });
              var c2 = Promise.resolve();
              rep.forEach(function (r) {
                c2 = c2.then(function () {
                  r.scopeGroupFilter = null;
                  if (!r.scopeGroups || !r.scopeGroups.length) {
                    r.scopeGroups = [{ id: fallbackGroupId }];
                  }
                  return rpc('Set', { typeName: 'CustomReportSchedule', entity: r });
                });
              });
              return c2;
            });
          });
        });
      });
      return chain;
    });
  }

  function reportUnclearable(addInBlockers, manual, name) {
    log('"' + name + '" could not be deleted - provider action required.', 'err');
    if (addInBlockers.length) {
      log('Blocked by marketplace add-in(s). Ask your provider to disable auto-enrollment in MyAdmin (or remove this group from their scope):', 'err');
      var chain = Promise.resolve();
      addInBlockers.forEach(function (a) {
        chain = chain.then(function () {
          return getEntities('AddIn', { id: a.id }).then(function (rows) {
            var full = rows && rows[0];
            var nm = a.id;
            var auto = '?';
            if (full) {
              nm = (full.configuration && full.configuration.name) || a.id;
              auto = full.isAutoEnrollEnabled ? 'ON' : 'OFF';
            }
            log('  - ' + nm + '   (auto-enroll: ' + auto + ',  id: ' + a.id + ')', 'err');
            return null;
          }).catch(function () {
            log('  - Add-In ' + a.id, 'err');
            return null;
          });
        });
      });
      chain.then(function () {
        log('Once disabled in MyAdmin, click Refresh and try Delete again.', 'info');
      });
    }
    if (manual.length) {
      log('Other blockers needing manual review: ' + manual.join(', '), 'err');
    }
  }

  // ---------------------------- lifecycle ---------------------------------
  return {
    initialize: function (freshApi, freshState, initializeCallback) {
      api = freshApi;
      injectStyles();
      elRoot = document.getElementById('cdg-root');
      buildUI();
      try {
        api.getSession(function (session) {
          creds = (session && session.credentials) ? session.credentials : session;
          initializeCallback();
        });
      } catch (e) {
        initializeCallback();
      }
    },
    focus: function (freshApi, freshState) {
      api = freshApi;
      injectStyles();
      if (!creds) {
        try {
          api.getSession(function (session) {
            creds = (session && session.credentials) ? session.credentials : session;
            loadGroups();
          });
          return;
        } catch (e) { /* fallthrough */ }
      }
      loadGroups();
    },
    blur: function () { }
  };
};
