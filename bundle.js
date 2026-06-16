/* ==========================================================================
 *  Dynasty Group Delete Solutions — v2.0
 *  MyGeotab Add-In (vanilla JS, CSP / Trusted Types compliant)
 *
 *  Aaj ki saari logic is mein hai:
 *   - Group delete ke liye "scan" par bharosa NAHI; Remove try karke
 *     GroupRelationViolatedException ka data parse karta hai (16 categories).
 *   - Jo API se clear ho sakta hai (Device/User/Zone/Rule/Driver/eventRules/
 *     dvirLogs/defects/inspectionTemplates/maintenanceWorkJobs + groupFilters +
 *     customReportSchedules) usay khud clear karta hai aur loop kar ke delete.
 *   - Jo signed/auto-enroll marketplace add-ins block karein (API se nahi
 *     hatte) unka EXACT naam + auto-enroll status dikhata hai:
 *     "MyAdmin se in ka auto-enroll off karein".
 *
 *  CSP rules: koi innerHTML-with-HTML nahi (createElement/textContent),
 *  saare variables function scope ke andar (var), onclick attribute nahi
 *  (addEventListener), JS strings mein literal newline nahi.
 * ========================================================================== */
geotab.addin.dynastyGroupDelete = function () {
  'use strict';

  var api = null;
  var creds = null;
  var elRoot = null;
  var elSelect = null;
  var elDeleteBtn = null;
  var elRefreshBtn = null;
  var elLog = null;
  var elStatus = null;
  var busy = false;

  // Built-in / system groups jo delete nahi hote
  var SYSTEM_PREFIX = 'Group';
  var COMPANY_ID = 'GroupCompanyId';

  // category -> [typeName, [fields]] jo API se clear ho sakte hain
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
  function setBusy(state) {
    busy = state;
    if (elDeleteBtn) { elDeleteBtn.disabled = state; }
    if (elSelect) { elSelect.disabled = state; }
    if (elRefreshBtn) { elRefreshBtn.disabled = state; }
  }

  function setStatus(text, kind) {
    if (!elStatus) { return; }
    elStatus.textContent = text || '';
    elStatus.className = 'dgd-status' + (kind ? ' dgd-' + kind : '');
  }

  function log(msg, kind) {
    if (!elLog) { return; }
    var row = document.createElement('div');
    row.className = 'dgd-logrow' + (kind ? ' dgd-' + kind : '');
    var time = document.createElement('span');
    time.className = 'dgd-logtime';
    time.textContent = new Date().toLocaleTimeString();
    var text = document.createElement('span');
    text.textContent = msg;
    row.appendChild(time);
    row.appendChild(text);
    elLog.appendChild(row);
    elLog.scrollTop = elLog.scrollHeight;
  }

  function clearLog() {
    if (!elLog) { return; }
    while (elLog.firstChild) { elLog.removeChild(elLog.firstChild); }
  }

  function buildUI() {
    while (elRoot.firstChild) { elRoot.removeChild(elRoot.firstChild); }

    var title = document.createElement('h2');
    title.className = 'dgd-title';
    title.textContent = 'Group Delete';
    elRoot.appendChild(title);

    var sub = document.createElement('p');
    sub.className = 'dgd-sub';
    sub.textContent = 'Select a group to delete. References (devices, users, zones, rules, report filters) are cleared automatically. Marketplace add-ins that block deletion are listed for your provider to disable.';
    elRoot.appendChild(sub);

    var bar = document.createElement('div');
    bar.className = 'dgd-bar';

    elSelect = document.createElement('select');
    elSelect.className = 'dgd-select';
    var ph = document.createElement('option');
    ph.value = '';
    ph.textContent = 'Loading groups...';
    elSelect.appendChild(ph);
    bar.appendChild(elSelect);

    elDeleteBtn = document.createElement('button');
    elDeleteBtn.className = 'dgd-btn dgd-btn-danger';
    elDeleteBtn.type = 'button';
    elDeleteBtn.textContent = 'Delete group';
    elDeleteBtn.addEventListener('click', onDeleteClick);
    bar.appendChild(elDeleteBtn);

    elRefreshBtn = document.createElement('button');
    elRefreshBtn.className = 'dgd-btn';
    elRefreshBtn.type = 'button';
    elRefreshBtn.textContent = 'Refresh';
    elRefreshBtn.addEventListener('click', loadGroups);
    bar.appendChild(elRefreshBtn);

    elRoot.appendChild(bar);

    elStatus = document.createElement('div');
    elStatus.className = 'dgd-status';
    elRoot.appendChild(elStatus);

    var logTitle = document.createElement('div');
    logTitle.className = 'dgd-logtitle';
    logTitle.textContent = 'Execution log';
    elRoot.appendChild(logTitle);

    elLog = document.createElement('div');
    elLog.className = 'dgd-log';
    elRoot.appendChild(elLog);
  }

  function loadGroups() {
    if (busy) { return; }
    setStatus('Loading groups...', 'info');
    getEntities('Group').then(function (groups) {
      while (elSelect.firstChild) { elSelect.removeChild(elSelect.firstChild); }
      var ph = document.createElement('option');
      ph.value = '';
      ph.textContent = '-- Select a group --';
      elSelect.appendChild(ph);

      groups.filter(function (g) {
        return !isSystemGroup(g.id);
      }).sort(function (a, b) {
        return String(a.name || '').localeCompare(String(b.name || ''));
      }).forEach(function (g) {
        var opt = document.createElement('option');
        opt.value = g.id;
        opt.textContent = (g.name || g.id);
        elSelect.appendChild(opt);
      });
      setStatus('Ready. ' + (elSelect.options.length - 1) + ' deletable groups.', 'ok');
    }).catch(function (e) {
      setStatus('Could not load groups: ' + e.message, 'err');
    });
  }

  function onDeleteClick() {
    var gid = elSelect.value;
    if (!gid) { setStatus('Pehle ek group select karein.', 'warn'); return; }
    var name = elSelect.options[elSelect.selectedIndex].textContent;
    clearLog();
    setBusy(true);
    setStatus('Deleting "' + name + '"...', 'info');
    log('Starting delete for "' + name + '" (' + gid + ')');
    runDelete(gid, name, 1);
  }

  // ----------------------- core delete (loop) -----------------------------
  function runDelete(gid, name, iter) {
    var MAX_ITERS = 8;
    if (iter > MAX_ITERS) {
      setStatus('Stopped: zyada iterations. Baqi blockers manual handling maangte hain.', 'warn');
      setBusy(false);
      return;
    }

    rpc('Remove', { typeName: 'Group', entity: { id: gid } }).then(function () {
      log('Group deleted successfully.', 'ok');
      setStatus('"' + name + '" delete ho gaya.', 'ok');
      setBusy(false);
      loadGroups();
    }).catch(function (err) {
      if (err.geotabName !== 'GroupRelationViolatedException' || !err.dataInfo) {
        log('Error: ' + err.message, 'err');
        setStatus('Delete fail: ' + err.message, 'err');
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
    var tasks = [];        // promises jo clear karenge
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
      log('No clearable references found but delete still blocked.', 'warn');
      setBusy(false);
      return;
    }

    if (tasks.length === 0) {
      // Sirf add-ins / manual bache — clear karne ko kuch nahi
      reportUnclearable(addInBlockers, manual, name);
      setBusy(false);
      return;
    }

    log('Iteration ' + iter + ': clearing ' + tasks.length + ' reference(s)...');
    Promise.all(tasks).then(function () {
      runDelete(gid, name, iter + 1);
    }).catch(function (e) {
      log('Clear step error: ' + e.message, 'err');
      setBusy(false);
    });
  }

  // entity ke groups/companyGroups/driverGroups se group nikaal kar Set
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
      log('  cleared ' + typeName + ' "' + (ent.name || id) + '"');
      return rpc('Set', { typeName: typeName, entity: ent });
    });
  }

  // GroupFilter se group ki condition nikaalna (1 bache to single-group flatten)
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
            log('  cleared report group-filter ' + f.id);
            return rpc('Set', { typeName: 'GroupFilter', entity: gf }).catch(function () {
              // fallback: us filter ko use karne wale report schedule ko re-scope
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

  // Add-ins jo API se clear nahi hote — naam + auto-enroll status dikhana
  function reportUnclearable(addInBlockers, manual, name) {
    setStatus('"' + name + '" delete nahi ho saka — provider action chahiye.', 'warn');
    if (addInBlockers.length) {
      log('Blocked by marketplace add-in(s). MyAdmin se in ka auto-enroll OFF karwayein (ya group scope se hatayein):', 'warn');
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
            log('  • ' + nm + '   (auto-enroll: ' + auto + ',  id: ' + a.id + ')', 'warn');
            return null;
          }).catch(function () {
            log('  • Add-In ' + a.id, 'warn');
            return null;
          });
        });
      });
      chain.then(function () {
        log('Apne provider (Dynasty) se request karein, ya MyAdmin access ho to wahan se off karein, phir Refresh + dobara Delete.', 'info');
      });
    }
    if (manual.length) {
      log('Other blockers needing manual review: ' + manual.join(', '), 'warn');
    }
  }

  // ---------------------------- lifecycle ---------------------------------
  return {
    initialize: function (freshApi, freshState, initializeCallback) {
      api = freshApi;
      elRoot = document.getElementById('cdg-root');
      buildUI();
      // session credentials (direct apiv1 calls ke liye, taake error.data mile)
      try {
        api.getSession(function (session) {
          creds = (session && session.credentials) ? session.credentials : session;
          initializeCallback();
        });
      } catch (e) {
        // agar getSession promise based ho
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
