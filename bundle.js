/* ==========================================================================
 *  Dynasty Group Delete Solutions - v2.0
 *  MyGeotab Add-In (vanilla JS, CSP / Trusted Types compliant)
 *
 *  Uses the existing styles.css classes (cdg-* + Zenith design tokens),
 *  so styles.css does NOT need to change.
 *
 *  Approach:
 *   - Does NOT rely on scanning entity types. It attempts Remove and parses
 *     the GroupRelationViolatedException data payload (16 relation categories).
 *   - Auto-clears everything the API can clear (Device/User/Zone/Rule/Driver/
 *     eventRules/dvirLogs/defects/inspectionTemplates/maintenanceWorkJobs +
 *     groupFilters + customReportSchedules), then loops until the group deletes.
 *   - For signed / auto-enroll marketplace add-ins that cannot be changed via
 *     the API, it shows the exact add-in name and auto-enroll status with
 *     guidance to disable auto-enrollment in MyAdmin.
 *
 *  CSP rules: no innerHTML with HTML strings (createElement/textContent only),
 *  all variables in function scope (var), no onclick attributes
 *  (addEventListener), no literal newlines inside JS strings.
 * ========================================================================== */
geotab.addin.dynastyGroupDelete = function () {
  'use strict';

  var api = null;
  var creds = null;
  var elRoot = null;
  var elSelect = null;
  var elConfirm = null;
  var elDeleteBtn = null;
  var elRefreshBtn = null;
  var elLog = null;
  var busy = false;

  // Built-in / system groups that cannot be deleted
  var SYSTEM_PREFIX = 'Group';
  var COMPANY_ID = 'GroupCompanyId';

  // category -> [typeName, [fields]] that can be cleared via the API
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
  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) { node.className = className; }
    if (text != null) { node.textContent = text; }
    return node;
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
    var line = el('div', 'cdg-log-line' + (kind ? ' cdg-log-' + kind : ''));
    line.textContent = '[' + new Date().toLocaleTimeString() + ']  ' + msg;
    elLog.appendChild(line);
    elLog.scrollTop = elLog.scrollHeight;
  }

  function clearLog() {
    if (!elLog) { return; }
    while (elLog.firstChild) { elLog.removeChild(elLog.firstChild); }
  }

  function buildUI() {
    while (elRoot.firstChild) { elRoot.removeChild(elRoot.firstChild); }

    var app = el('div', 'cdg-app');
    var content = el('div', 'cdg-content');

    var title = el('h2', 'heading-01-desktop', 'Group Delete');
    content.appendChild(title);

    var sub = el('p', 'body-01', 'Select a group to delete. References (devices, users, zones, rules, report filters) are cleared automatically. Marketplace add-ins that block deletion are listed for your provider to disable.');
    sub.style.color = 'var(--text-secondary)';
    content.appendChild(sub);

    // selector + actions row
    var row = el('div', 'cdg-row');

    var selWrap = el('div', 'cdg-select-wrap');
    var selLabel = el('div', 'zen-field-label__text', 'Group');
    selLabel.style.marginBottom = '4px';
    elSelect = el('select', 'zen-text-input');
    elSelect.style.width = '100%';
    var ph = el('option', null, 'Loading groups...');
    ph.value = '';
    elSelect.appendChild(ph);
    elSelect.addEventListener('change', syncDeleteEnabled);
    selWrap.appendChild(selLabel);
    selWrap.appendChild(elSelect);
    row.appendChild(selWrap);

    var actions = el('div', 'cdg-actions');
    elRefreshBtn = el('button', 'zen-button', 'Refresh');
    elRefreshBtn.type = 'button';
    elRefreshBtn.addEventListener('click', loadGroups);
    elDeleteBtn = el('button', 'zen-button zen-button--destructive', 'Delete group');
    elDeleteBtn.type = 'button';
    elDeleteBtn.disabled = true;
    elDeleteBtn.addEventListener('click', onDeleteClick);
    actions.appendChild(elRefreshBtn);
    actions.appendChild(elDeleteBtn);
    row.appendChild(actions);

    content.appendChild(row);

    // confirm group
    var confirmGroup = el('div', 'cdg-confirm-group');
    var confirmLabel = el('label', 'zen-checkbox__label');
    confirmLabel.style.display = 'flex';
    confirmLabel.style.alignItems = 'center';
    confirmLabel.style.gap = '8px';
    confirmLabel.style.cursor = 'pointer';
    elConfirm = el('input');
    elConfirm.type = 'checkbox';
    elConfirm.addEventListener('change', syncDeleteEnabled);
    var confirmText = el('span', 'body-04', 'I understand this permanently deletes the selected group and removes it from all referencing entities.');
    confirmLabel.appendChild(elConfirm);
    confirmLabel.appendChild(confirmText);
    confirmGroup.appendChild(confirmLabel);
    content.appendChild(confirmGroup);

    // execution log
    var logTitle = el('div', 'heading-05', 'Execution log');
    logTitle.style.marginTop = '8px';
    content.appendChild(logTitle);

    elLog = el('div', 'cdg-log');
    content.appendChild(elLog);

    app.appendChild(content);
    elRoot.appendChild(app);
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
    var tasks = [];        // promises that clear references
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
      // Only add-ins / manual items remain - nothing left to clear via API
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

  // Remove the group from an entity's groups/companyGroups/driverGroups and Set
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

  // Remove the group's condition from a GroupFilter (flatten to single group if one remains)
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
              // fallback: re-scope the report schedule that uses this filter
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

  // Add-ins that cannot be cleared via the API - show name + auto-enroll status
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
      elRoot = document.getElementById('cdg-root');
      buildUI();
      // session credentials (for direct apiv1 calls so we can read error.data)
      try {
        api.getSession(function (session) {
          creds = (session && session.credentials) ? session.credentials : session;
          initializeCallback();
        });
      } catch (e) {
        // in case getSession is promise based
        initializeCallback();
      }
    },
    focus: function (freshApi, freshState) {
      api = freshApi;
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
