/* ==========================================================================
 *  Dynasty Group Delete Solutions - v3.0 (Dynasty Safety Console "NOC" theme)
 *  MyGeotab Add-In (vanilla JS, CSP / Trusted Types compliant)
 *
 *  - Functionality is identical to the working v2.x build: attempt Remove,
 *    parse the GroupRelationViolatedException data payload (16 relation
 *    categories), auto-clear everything the API can clear, loop until the
 *    group deletes, and report unclearable marketplace add-ins with their
 *    auto-enroll status + MyAdmin guidance.
 *  - Front-end follows the Dynasty Safety Console design language: a
 *    self-contained dark "NOC" theme with a light-mode toggle, Dynasty amber
 *    branding, card panel, and a console-style execution log. All styling is
 *    injected at runtime via injectStyles() (no external styles.css), so it
 *    renders correctly inside MyGeotab's add-in iframe.
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
  var elConfirm = null;
  var elDeleteBtn = null;
  var elRefreshBtn = null;
  var elLog = null;
  var busy = false;

  var SYSTEM_PREFIX = 'Group';
  var COMPANY_ID = 'GroupCompanyId';

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

  // ---------------------------- styling -----------------------------------
  // Self-contained NOC theme. Each entry is a single line (no literal newlines
  // inside strings); joined and injected once via createElement/textContent.
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
      // header / brand
      '.cdg-header{display:flex;align-items:center;gap:14px;margin-bottom:18px}',
      '.cdg-shield{width:34px;height:34px;flex:0 0 auto}',
      '.cdg-brandtext{display:flex;flex-direction:column;gap:2px;flex:1 1 auto;min-width:0}',
      '.cdg-title{font-size:20px;font-weight:600;line-height:24px;margin:0;color:var(--text);letter-spacing:.2px}',
      '.cdg-eyebrow{font-size:11px;font-weight:600;letter-spacing:1.4px;text-transform:uppercase;color:var(--amber);margin:0}',
      '.cdg-theme{flex:0 0 auto;border:1px solid var(--border);background:var(--panel-2);color:var(--muted);font-family:var(--font);font-size:12px;font-weight:500;padding:6px 12px;border-radius:999px;cursor:pointer;outline:none}',
      '.cdg-theme:hover{color:var(--text);border-color:var(--amber)}',
      // card
      '.cdg-card{background:var(--panel);border:1px solid var(--border);border-radius:var(--radius);padding:20px;max-width:760px;box-shadow:0 1px 0 rgba(0,0,0,.25)}',
      '.cdg-intro{font-size:13px;line-height:20px;color:var(--muted);margin:0 0 16px}',
      // row + field + actions
      '.cdg-row{display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap}',
      '.cdg-field{flex:1 1 320px;min-width:240px}',
      '.cdg-label{display:block;font-size:11px;font-weight:600;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin-bottom:6px}',
      '.cdg-select{width:100%;height:38px;padding:8px 12px;font-family:var(--font);font-size:13px;color:var(--text);background:var(--panel-2);border:1px solid var(--border);border-radius:8px;outline:none;cursor:pointer;-webkit-appearance:none;-moz-appearance:none;appearance:none}',
      '.cdg-select:hover{border-color:var(--amber)}',
      '.cdg-select:focus{border-color:var(--amber);box-shadow:0 0 0 2px rgba(255,180,84,.18)}',
      '.cdg-actions{display:flex;gap:8px;flex-wrap:wrap}',
      // buttons
      '.cdg-btn{display:inline-flex;align-items:center;justify-content:center;height:38px;padding:0 16px;font-family:var(--font);font-size:13px;font-weight:600;color:var(--text);background:var(--panel-2);border:1px solid var(--border);border-radius:8px;cursor:pointer;outline:none}',
      '.cdg-btn:hover{border-color:var(--amber);color:var(--amber)}',
      '.cdg-btn:disabled,.cdg-btn[disabled]{opacity:.45;cursor:default;pointer-events:none}',
      '.cdg-btn-danger{color:#fff;background:var(--red);border-color:var(--red)}',
      '.cdg-btn-danger:hover{color:#fff;background:var(--red-strong);border-color:var(--red-strong)}',
      // confirm
      '.cdg-confirm{display:flex;gap:10px;align-items:flex-start;margin:16px 0 4px;padding:12px 14px;background:rgba(255,93,93,.08);border:1px solid rgba(255,93,93,.35);border-radius:8px}',
      '.cdg-check{position:relative;flex:0 0 auto;width:18px;height:18px;margin-top:1px}',
      '.cdg-check input{position:absolute;opacity:0;width:18px;height:18px;margin:0;cursor:pointer}',
      '.cdg-box{width:18px;height:18px;border:1px solid var(--border);border-radius:4px;background:var(--panel-2);display:flex;align-items:center;justify-content:center}',
      '.cdg-check input:checked + .cdg-box{background:var(--red);border-color:var(--red)}',
      '.cdg-box svg{display:none;width:12px;height:12px}',
      '.cdg-check input:checked + .cdg-box svg{display:block}',
      '.cdg-check input:focus + .cdg-box{box-shadow:0 0 0 2px rgba(255,93,93,.3)}',
      '.cdg-confirm-text{font-size:12px;line-height:18px;color:var(--text);cursor:pointer}',
      '.cdg-confirm-text b{color:var(--red)}',
      // log
      '.cdg-logtitle{font-size:11px;font-weight:600;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);margin:18px 0 8px}',
      '.cdg-log{background:var(--log-bg);border:1px solid var(--border);color:#cdd6e0;font-family:var(--mono);font-size:12px;line-height:1.7;padding:12px 14px;border-radius:8px;max-height:320px;overflow-y:auto}',
      '.cdg-log-empty{color:#5b6673;font-style:italic}',
      '.cdg-log-line{white-space:pre-wrap;word-break:break-word}',
      '.cdg-log-ok{color:var(--green)}',
      '.cdg-log-err{color:var(--red)}',
      '.cdg-log-info{color:#8fb9e8}',
      // footer
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
    var shield = svgEl('path', {
      d: 'M16 2 L28 6 V15 C28 22.5 22.8 27.6 16 30 C9.2 27.6 4 22.5 4 15 V6 Z',
      fill: 'var(--amber)'
    });
    var inner = svgEl('path', {
      d: 'M11 16 l3.4 3.4 L21 12.8',
      fill: 'none', stroke: '#0d1117', 'stroke-width': '2.6',
      'stroke-linecap': 'round', 'stroke-linejoin': 'round'
    });
    svg.appendChild(shield);
    svg.appendChild(inner);
    return svg;
  }

  function checkMark() {
    var svg = svgEl('svg', { viewBox: '0 0 16 16' });
    var path = svgEl('path', {
      d: 'M6.5 11.5 L3 8 l1.4 -1.4 L6.5 8.7 l5.1 -5.1 L13 4 z',
      fill: '#ffffff'
    });
    svg.appendChild(path);
    return svg;
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

  function syncDeleteEnabled() {
    if (!elDeleteBtn) { return; }
    var ok = !busy && elConfirm && elConfirm.checked && elSelect && elSelect.value;
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

  function buildUI() {
    while (elRoot.firstChild) { elRoot.removeChild(elRoot.firstChild); }

    elApp = el('div', 'cdg-app');

    // header
    var header = el('div', 'cdg-header');
    header.appendChild(shieldIcon());
    var brandText = el('div', 'cdg-brandtext');
    brandText.appendChild(el('div', 'cdg-eyebrow', 'Dynasty Group Delete'));
    brandText.appendChild(el('h1', 'cdg-title', 'Cascade Delete'));
    header.appendChild(brandText);
    elThemeBtn = el('button', 'cdg-theme', 'Light');
    elThemeBtn.type = 'button';
    elThemeBtn.addEventListener('click', toggleTheme);
    header.appendChild(elThemeBtn);
    elApp.appendChild(header);

    // card
    var card = el('div', 'cdg-card');
    card.appendChild(el('p', 'cdg-intro', 'Select a group to delete. Referencing entities (devices, users, zones, rules, report filters) are cleared automatically, then the group is removed. Marketplace add-ins that block deletion are listed for your provider to disable in MyAdmin.'));

    var row = el('div', 'cdg-row');
    var field = el('div', 'cdg-field');
    field.appendChild(el('label', 'cdg-label', 'Group'));
    elSelect = el('select', 'cdg-select');
    var ph = el('option', null, 'Loading groups...');
    ph.value = '';
    elSelect.appendChild(ph);
    elSelect.addEventListener('change', syncDeleteEnabled);
    field.appendChild(elSelect);
    row.appendChild(field);

    var actions = el('div', 'cdg-actions');
    elRefreshBtn = el('button', 'cdg-btn', 'Refresh');
    elRefreshBtn.type = 'button';
    elRefreshBtn.addEventListener('click', loadGroups);
    elDeleteBtn = el('button', 'cdg-btn cdg-btn-danger', 'Delete group');
    elDeleteBtn.type = 'button';
    elDeleteBtn.disabled = true;
    elDeleteBtn.addEventListener('click', onDeleteClick);
    actions.appendChild(elRefreshBtn);
    actions.appendChild(elDeleteBtn);
    row.appendChild(actions);
    card.appendChild(row);

    // confirm
    var confirm = el('div', 'cdg-confirm');
    var checkWrap = el('label', 'cdg-check');
    elConfirm = el('input');
    elConfirm.type = 'checkbox';
    elConfirm.id = 'cdg-confirm-input';
    elConfirm.addEventListener('change', syncDeleteEnabled);
    var box = el('span', 'cdg-box');
    box.appendChild(checkMark());
    checkWrap.appendChild(elConfirm);
    checkWrap.appendChild(box);
    var confirmLabel = el('label', 'cdg-confirm-text');
    confirmLabel.setAttribute('for', 'cdg-confirm-input');
    var warnBold = el('b', null, 'Cascade delete is irreversible.');
    confirmLabel.appendChild(warnBold);
    confirmLabel.appendChild(document.createTextNode(' This permanently deletes the selected group and removes it from all referencing entities.'));
    confirm.appendChild(checkWrap);
    confirm.appendChild(confirmLabel);
    card.appendChild(confirm);

    // log
    card.appendChild(el('div', 'cdg-logtitle', 'Execution log'));
    elLog = el('div', 'cdg-log');
    card.appendChild(elLog);

    elApp.appendChild(card);

    // footer
    elApp.appendChild(el('div', 'cdg-footer', 'Dynasty Communications'));

    elRoot.appendChild(elApp);
    logPlaceholder();
  }

  function loadGroups() {
    if (busy) { return; }
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

  function onDeleteClick() {
    var gid = elSelect.value;
    if (!gid) { log('Please select a group first.', 'err'); return; }
    if (!elConfirm.checked) { log('Please confirm before deleting.', 'err'); return; }
    var name = elSelect.options[elSelect.selectedIndex].textContent;
    clearLog();
    setBusy(true);
    log('Starting delete for "' + name + '" (' + gid + ')', 'info');
    runDelete(gid, name, 1);
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
      if (elConfirm) { elConfirm.checked = false; }
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
    var k, cat;

    for (k in info) {
      if (!info.hasOwnProperty(k) || k === 'group') { continue; }
      cat = info[k];
      if (!Array.isArray(cat) || cat.length === 0) { continue; }

      if (k === 'addIns') {
        addInBlockers = cat.slice();
      } else if (k === 'groupFilters') {
        tasks.push(clearGroupFilters(cat, gid));
      } else if (CAT_MAP[k]) {
        cat.forEach(function (item) {
          tasks.push(clearEntity(CAT_MAP[k][0], CAT_MAP[k][1], item.id, gid));
        });
      } else {
        manual.push(k + ' (' + cat.length + ')');
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

    log('Iteration ' + iter + ': clearing ' + tasks.length + ' reference(s)...', 'info');
    Promise.all(tasks).then(function () {
      runDelete(gid, name, iter + 1);
    }).catch(function (e) {
      log('Clear step error: ' + e.message, 'err');
      setBusy(false);
    });
  }

  function clearEntity(typeName, fields, id, gid) {
    return getEntities(typeName, { id: id }).then(function (rows) {
      var ent = rows && rows[0];
      if (!ent) { return null; }
      var changed = false;
      fields.forEach(function (f) {
        if (Array.isArray(ent[f]) && hasGroup(ent[f], gid)) {
          ent[f] = ent[f].filter(function (g) { return g.id !== gid; });
          changed = true;
        }
      });
      if (!changed) { return null; }
      log('  cleared ' + typeName + ' "' + (ent.name || id) + '"', 'info');
      return rpc('Set', { typeName: typeName, entity: ent });
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
                    r.scopeGroups = [{ id: COMPANY_ID }];
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
