/* EA Inbox Management. Shows EA's triage for the selected mail with one-click
   actions, and tags the mail with a "Notion: <lane>" category (and optional
   note) so EA can turn it into a Notion card later.
   Read mode talks to Microsoft Graph with a token from Nested App Authentication
   (NAA), so everything lands on the server directly. Classic Outlook neither
   syncs Office custom properties nor allows EWS tokens any more.
   Compose mode (no server id yet) keeps using Office.js locally. */
(function () {
  "use strict";

  var LANES = [
    // Same order, names and colours as Leon's Notion board.
    { key: "backlog", label: "Backlog",     category: "Notion: Backlog",             color: "Preset12", tone: "default" },
    { key: "week",    label: "Sprint 1W",   category: "Notion: Sprint 1 week",       color: "Preset7",  tone: "blue" },
    { key: "nextday", label: "Next day",    category: "Notion: Sprint next day",     color: "Preset7",  tone: "blue" },
    { key: "today",   label: "Sprint 1D",   category: "Notion: Sprint 1 day",        color: "Preset0",  tone: "red" },
    { key: "onhold",  label: "Waiting for", category: "Notion: On hold/waiting for", color: "Preset7",  tone: "blue" }
  ];
  var NOTE_KEY = "evaNote";
  var LANE_KEY = "evaLane";
  var TRIAGE_KEY = "evaTriage";
  var ACTION_KEY = "evaAction";
  var ACTION_CAT = { displayName: "EA: actie", color: "Preset8" };
  var busy = false;

  var $ = function (id) { return document.getElementById(id); };

  // ---- Office.js callback -> Promise helpers ------------------------------
  function call(fn) {
    return new Promise(function (resolve, reject) {
      fn(function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
        else reject(res.error || new Error("Onbekende fout"));
      });
    });
  }
  function item() { return Office.context.mailbox.item; }
  function readMode() { return !!(item() && item().itemId); }

  function laneByCategory(name) {
    for (var i = 0; i < LANES.length; i++) if (LANES[i].category === name) return LANES[i];
    return null;
  }
  function msg(err) {
    if (!err) return "onbekende fout";
    var code = err.errorCode || err.code;
    var text = err.message || err.name || "onbekende fout";
    return code && text.indexOf(code) < 0 ? text + " (" + code + ")" : text;
  }

  // ---- Graph via Nested App Authentication --------------------------------
  var AUTH = {
    clientId: "cee9ce67-5476-42a8-b9cb-46d110f60d08",
    authority: "https://login.microsoftonline.com/650e65a6-7112-4147-9ef8-c7e6ca69613c",
    scopes: ["https://graph.microsoft.com/Mail.ReadWrite"]
  };
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var PROP_ID = "String {00020329-0002-0000-C000-000000000046} Name cecp-2c2658f7-4faa-450c-9387-4f9ccf079dbb";
  var pcaPromise = null;

  function naaSupported() {
    try { return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"); } catch (e) { return false; }
  }
  function getPca() {
    if (!pcaPromise) {
      if (typeof msal === "undefined" || !msal.createNestablePublicClientApplication) {
        return Promise.reject(new Error("Aanmeldbibliotheek niet geladen"));
      }
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: { clientId: AUTH.clientId, authority: AUTH.authority },
        cache: { cacheLocation: "localStorage" }
      }).catch(function (e) { pcaPromise = null; throw e; });
    }
    return pcaPromise;
  }
  function loginHint() {
    try { return Office.context.mailbox.userProfile.emailAddress || undefined; } catch (e) { return undefined; }
  }
  // interactive: false = only silent (used on pane load, where a popup without a click can be blocked).
  function getGraphToken(interactive) {
    if (!naaSupported()) {
      var e = new Error("Deze Outlook-versie ondersteunt geen aanmelden via de add-in (Nested App Authentication). Werk Outlook bij of gebruik Outlook op het web.");
      e.code = "NAA_UNSUPPORTED";
      return Promise.reject(e);
    }
    return getPca().then(function (pca) {
      var account = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
      var req = { scopes: AUTH.scopes };
      if (account) req.account = account; else req.loginHint = loginHint();
      return pca.acquireTokenSilent(req).catch(function (err) {
        if (interactive === false) {
          var e2 = new Error("Aanmelden nodig");
          e2.code = "LOGIN_REQUIRED";
          e2.cause = err;
          throw e2;
        }
        return pca.acquireTokenPopup({ scopes: AUTH.scopes, loginHint: loginHint() });
      }).then(function (res) {
        if (res.account) pca.setActiveAccount(res.account);
        return res.accessToken;
      });
    });
  }

  function graph(method, path, body, interactive) {
    return getGraphToken(interactive).then(function (token) {
      var opts = { method: method, headers: { Authorization: "Bearer " + token } };
      if (body !== undefined) {
        opts.headers["Content-Type"] = "application/json";
        opts.body = JSON.stringify(body);
      }
      return fetch(path.indexOf("https://") === 0 ? path : GRAPH + path, opts);
    }).then(function (r) {
      if (r.status === 204) return null;
      return r.text().then(function (txt) {
        var data = null;
        try { data = txt ? JSON.parse(txt) : null; } catch (e) { data = null; }
        if (!r.ok) {
          var ge = (data && data.error) || {};
          var err = new Error("Graph " + r.status + (ge.message ? ": " + ge.message : ""));
          err.code = ge.code || ("HTTP" + r.status);
          throw err;
        }
        return data;
      });
    });
  }

  function restId() {
    return Office.context.mailbox.convertToRestId(item().itemId, Office.MailboxEnums.RestVersion.v2_0);
  }
  function msgPath(id) { return "/me/messages/" + encodeURIComponent(id); }

  function getMessage(id, interactive) {
    var q = "?$select=id,categories,isRead&$expand=singleValueExtendedProperties($filter=id eq '" + PROP_ID + "')";
    return graph("GET", msgPath(id) + q, undefined, interactive).then(function (m) {
      var props = {};
      var ext = (m && m.singleValueExtendedProperties) || [];
      if (ext.length && ext[0].value) {
        try { props = JSON.parse(ext[0].value) || {}; } catch (e) { props = {}; }
      }
      return { id: m.id, categories: m.categories || [], isRead: !!m.isRead, props: props };
    });
  }

  // Read-modify-write of categories and the custom-property JSON in one PATCH.
  // Writes are queued so a debounced note save never interleaves with a lane click.
  var writeQueue = Promise.resolve();
  function mutateMessage(id, fn) {
    var run = function () {
      return getMessage(id).then(function (m) {
        var cats = m.categories.slice();
        var props = Object.assign({}, m.props);
        var res = fn(cats, props) || {};
        var patch = {};
        if (res.categories) patch.categories = res.categories;
        if (res.props) {
          // Office custom properties: a JSON object of string values.
          var clean = {};
          Object.keys(res.props).forEach(function (k) {
            var v = res.props[k];
            if (v === undefined || v === null || v === "") return;
            clean[k] = typeof v === "string" ? v : JSON.stringify(v);
          });
          patch.singleValueExtendedProperties = [{ id: PROP_ID, value: JSON.stringify(clean) }];
        }
        if (!patch.categories && !patch.singleValueExtendedProperties) return m;
        return graph("PATCH", msgPath(id), patch).then(function () { return m; });
      });
    };
    var p = writeQueue.then(run, run);
    writeQueue = p.catch(function () {});
    return p;
  }

  // ---- Master categories (local Office.js is fine for these) ---------------
  function getMasterCategories() {
    return call(function (cb) { Office.context.mailbox.masterCategories.getAsync(cb); }).then(function (v) { return v || []; });
  }
  function ensureMasterCategories(list) {
    list = list || LANES.map(function (l) { return { displayName: l.category, color: l.color }; });
    return getMasterCategories().then(function (existing) {
      var have = {};
      existing.forEach(function (c) { have[c.displayName.toLowerCase()] = true; });
      var missing = list.filter(function (m) { return !have[m.displayName.toLowerCase()]; });
      if (!missing.length) return;
      return call(function (cb) { Office.context.mailbox.masterCategories.addAsync(missing, cb); });
    }).catch(function () { /* colour only; the category itself is set via Graph */ });
  }

  // ---- Compose mode: no server id yet, so stay local via Office.js ---------
  function localCategories() {
    return call(function (cb) { item().categories.getAsync(cb); }).then(function (v) {
      return (v || []).map(function (c) { return c.displayName; });
    });
  }
  function localProps() { return call(function (cb) { item().loadCustomPropertiesAsync(cb); }); }
  function localApply(lane, note) {
    return localCategories().then(function (cats) {
      var others = cats.filter(function (n) { return laneByCategory(n) && (!lane || n !== lane.category); });
      var chain = others.length ? call(function (cb) { item().categories.removeAsync(others, cb); }) : Promise.resolve();
      return chain.then(function () {
        if (lane && cats.indexOf(lane.category) < 0) return call(function (cb) { item().categories.addAsync([lane.category], cb); });
      });
    }).then(localProps).then(function (props) {
      if (note) props.set(NOTE_KEY, note); else props.remove(NOTE_KEY);
      if (lane) props.set(LANE_KEY, lane.category); else props.remove(LANE_KEY);
      return call(function (cb) { props.saveAsync(cb); });
    });
  }

  // ---- Lanes and note ------------------------------------------------------
  // lane null = remove lane and note.
  function applyLane(lane, note) {
    if (!readMode()) return localApply(lane, note);
    return mutateMessage(restId(), function (cats, props) {
      var next = cats.filter(function (n) { return !laneByCategory(n); });
      if (lane) next.push(lane.category);
      props[NOTE_KEY] = lane ? note : "";
      props[LANE_KEY] = lane ? lane.category : "";
      return { categories: next, props: props };
    });
  }
  function saveNote(note, lane) {
    if (!readMode()) return localApply(lane, note);
    return mutateMessage(restId(), function (cats, props) {
      props[NOTE_KEY] = note;
      props[LANE_KEY] = lane ? lane.category : "";
      return { props: props };
    });
  }

  function pick(lane) {
    if (busy) return;
    var note = $("note").value.trim();
    setBusy(true);
    setStatus("Bezig...");
    ensureMasterCategories()
      .then(function () { return applyLane(lane, note); })
      .then(function () { showDone(lane, note); setStatus(""); })
      .catch(function (err) { setStatus("Dat lukte niet: " + msg(err), true); })
      .then(function () { setBusy(false); });
  }

  function undo() {
    if (busy) return;
    setBusy(true);
    setStatus("Ongedaan maken...");
    applyLane(null, "")
      .then(function () {
        showDone(null);
        $("note").value = "";
        setStatus("Verwijderd. EA pakt deze mail niet op.");
      })
      .catch(function (err) { setStatus("Ongedaan maken lukte niet: " + msg(err), true); })
      .then(function () { setBusy(false); });
  }

  // ---- EA triage -----------------------------------------------------------
  // EA writes "evaTriage" into the custom-property JSON of a mail (tools/mail_triage.py).
  // Archive and "gelezen" run directly via Graph; other clicks are recorded as
  // "evaAction" plus the category "EA: actie" for EA to execute.
  var folderCache = {};    // normalised path -> folder id
  var childCache = {};     // parent id ("" = top level) -> [{id, displayName}]
  var recentUndo = null;   // undo for a real move; survives a mail switch
  var stateUndo = null;    // undo for the state shown in the triage card
  var currentTriage = null;

  function laneFromTriage(v) {
    if (!v) return null;
    var s = String(v).toLowerCase().replace(/^notion:\s*/, "");
    for (var i = 0; i < LANES.length; i++) {
      var l = LANES[i];
      if (l.key === s || l.label.toLowerCase() === s || l.category.toLowerCase() === "notion: " + s) return l;
    }
    if (s.indexOf("on hold") === 0) return LANES.filter(function (l) { return l.key === "onhold"; })[0];
    return null;
  }

  function actionLabel(a) {
    if (a.label) return a.label;
    switch (a.type) {
      case "archive": return "Naar map " + (a.folderLabel || a.folderPath || "");
      case "notion":
        var l = laneFromTriage(a.lane);
        return "Naar Notion: " + (l ? l.label : (a.lane || ""));
      case "draft": return "Laat EA een reactie schrijven";
      case "fyi": return "Markeer als gelezen";
      default: return "Laten staan";
    }
  }

  function actionParams(a) {
    var p = {};
    Object.keys(a).forEach(function (k) { if (k !== "type" && k !== "label" && k !== "reason") p[k] = a[k]; });
    if (a.type === "archive") p.markRead = true;
    return p;
  }

  function parseJson(v) {
    if (!v) return null;
    if (typeof v === "object") return v;
    try { return JSON.parse(v); } catch (e) { return null; }
  }

  function formatAt(iso) {
    var d = iso ? new Date(iso) : null;
    if (!d || isNaN(d.getTime())) return "";
    var pad = function (n) { return (n < 10 ? "0" : "") + n; };
    return "triage " + d.getDate() + "/" + (d.getMonth() + 1) + " " + pad(d.getHours()) + ":" + pad(d.getMinutes());
  }

  // ---- Folders and moves via Graph -----------------------------------------
  function norm(s) { return String(s).toLowerCase().replace(/\s*\/\s*/g, "/").replace(/\s+/g, " ").trim(); }

  function listChildren(parentId) {
    if (childCache[parentId]) return Promise.resolve(childCache[parentId]);
    var first = parentId
      ? "/me/mailFolders/" + encodeURIComponent(parentId) + "/childFolders?$top=100&$select=id,displayName"
      : "/me/mailFolders?$top=100&$select=id,displayName";
    var out = [];
    function page(url) {
      return graph("GET", url).then(function (d) {
        out = out.concat((d && d.value) || []);
        if (d && d["@odata.nextLink"]) return page(d["@odata.nextLink"]);
        childCache[parentId] = out;
        return out;
      });
    }
    return page(first);
  }

  // Folder names can contain "/" themselves ("6. Advisors / Partners"), so at each
  // level try the longest run of path segments that matches a folder name.
  function findFolderId(path) {
    var key = norm(path);
    if (folderCache[key]) return Promise.resolve(folderCache[key]);
    var parts = String(path).split("/");
    function step(parentId, rest) {
      if (!rest.length) return Promise.resolve(parentId);
      return listChildren(parentId).then(function (kids) {
        for (var k = rest.length; k >= 1; k--) {
          var want = norm(rest.slice(0, k).join("/"));
          for (var i = 0; i < kids.length; i++) {
            if (norm(kids[i].displayName) === want) return step(kids[i].id, rest.slice(k));
          }
        }
        var e = new Error("Map niet gevonden: " + rest.join("/").trim());
        e.code = "FOLDER_NOT_FOUND";
        throw e;
      });
    }
    return step("", parts).then(function (id) { folderCache[key] = id; return id; });
  }

  function setRead(id, read) { return graph("PATCH", msgPath(id), { isRead: !!read }); }
  function moveTo(id, destinationId) {
    return graph("POST", msgPath(id) + "/move", { destinationId: destinationId }).then(function (m) { return m && m.id; });
  }

  // ---- Recorded actions ----------------------------------------------------
  function recordAction(a) {
    var payload = { type: a.type, params: actionParams(a), label: actionLabel(a), at: new Date().toISOString() };
    return ensureMasterCategories([ACTION_CAT]).then(function () {
      return mutateMessage(restId(), function (cats, props) {
        props[ACTION_KEY] = JSON.stringify(payload);
        if (cats.indexOf(ACTION_CAT.displayName) < 0) cats.push(ACTION_CAT.displayName);
        return { categories: cats, props: props };
      });
    }).then(function () { return payload; });
  }
  function clearRecordedAction() {
    return mutateMessage(restId(), function (cats, props) {
      delete props[ACTION_KEY];
      return { categories: cats.filter(function (n) { return n !== ACTION_CAT.displayName; }), props: props };
    });
  }

  function pendingText(a) {
    switch (a.type) {
      case "archive": return "EA verplaatst hem zo naar " + ((a.params && (a.params.folderLabel || a.params.folderPath)) || "de map");
      case "draft": return "Staat klaar voor EA: concept volgt";
      case "fyi": return "Staat klaar voor EA: wordt als gelezen gemarkeerd";
      default: return "Staat klaar voor EA: " + (a.label || a.type);
    }
  }

  // ---- Executing a click ---------------------------------------------------
  function runAction(a) {
    if (busy || !a) return;
    if (a.type === "notion") {
      var lane = laneFromTriage(a.lane);
      if (!lane) { setStatus("Onbekende lane in de triage: " + (a.lane || "leeg"), true); return; }
      if (a.cardTitle && !$("note").value.trim()) $("note").value = a.cardTitle;
      pick(lane);
      return;
    }
    if (a.type === "none" || !a.type) {
      showTriageState("Blijft staan", false);
      return;
    }
    var id = restId();
    setBusy(true);
    setStatus("Bezig...");
    var direct;
    if (a.type === "fyi") {
      direct = getMessage(id).then(function (m) {
        return setRead(id, true).then(function () {
          return { text: "Gemarkeerd als gelezen", undo: function () { return setRead(id, m.isRead); } };
        });
      });
    } else if (a.type === "archive") {
      direct = !a.folderPath ? Promise.reject(new Error("Geen map in de triage")) : getMessage(id).then(function (m) {
        // Resolve the folder first, so a missing folder changes nothing.
        return findFolderId(a.folderPath).then(function (fid) {
          return (m.isRead ? Promise.resolve() : setRead(id, true)).then(function () { return moveTo(id, fid); });
        }).then(function (newId) {
          var label = a.folderLabel || a.folderPath;
          return {
            moved: true,
            text: "Gearchiveerd in " + label,
            undo: function () {
              if (!newId) return Promise.reject(new Error("terugzetten kan niet, zoek hem in " + a.folderPath));
              return moveTo(newId, "inbox").then(function (backId) {
                if (!m.isRead && backId) return setRead(backId, false);
              });
            }
          };
        });
      });
    } else {
      direct = Promise.reject({ record: true });
    }
    direct
      .then(function (res) {
        if (res.moved) {
          showRecent(res.text, res.undo);
          showTriageState(res.text, false);
        } else {
          showTriageState(res.text, true, res.undo);
        }
        setStatus("");
      }, function (err) {
        // Draft and other actions, or a direct move that failed: record it for EA.
        return recordAction(a).then(function (payload) {
          showTriageState(pendingText(payload), true, clearRecordedAction);
          setStatus(err && !err.record ? "Direct uitvoeren lukte niet (" + msg(err) + "), EA doet het." : "");
        }, function (err2) { throw (err && !err.record) ? err : err2; });
      })
      .catch(function (err) { setStatus("Dat lukte niet: " + msg(err), true); })
      .then(function () { setBusy(false); });
  }

  // ---- Triage UI -----------------------------------------------------------
  function showTriageState(text, canUndo, undoFn) {
    $("triageActions").hidden = !!text;
    $("triageState").hidden = !text;
    $("triageStateText").textContent = text || "";
    stateUndo = canUndo ? (undoFn || null) : null;
    $("triageUndo").hidden = !stateUndo;
  }

  function showRecent(text, undoFn) {
    recentUndo = undoFn || null;
    $("recentText").textContent = text || "";
    $("recent").hidden = !text;
    $("recentUndo").hidden = !recentUndo;
  }

  function undoTriageState() {
    if (busy || !stateUndo) return;
    var fn = stateUndo;
    setBusy(true);
    setStatus("Ongedaan maken...");
    fn().then(function () {
      showTriageState(null);
      setStatus("Teruggezet.");
    }, function (err) { setStatus("Ongedaan maken lukte niet: " + msg(err), true); })
      .then(function () { setBusy(false); });
  }

  function undoRecent() {
    if (busy || !recentUndo) return;
    var fn = recentUndo;
    setBusy(true);
    setStatus("Terugzetten naar Postvak IN...");
    fn().then(function () {
      showRecent(null);
      setStatus("Terug in Postvak IN.");
    }, function (err) { setStatus("Terugzetten lukte niet: " + msg(err), true); })
      .then(function () { setBusy(false); });
  }

  function rich(s) {
    return String(s || "").replace(/[&<>]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]; }).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>");
  }

  function renderTriage(t, action) {
    currentTriage = t;
    var rm = readMode(); // triage only applies to a received mail
    $("triage").hidden = !t || !rm;
    $("noTriage").hidden = !!t || !rm;
    $("noTriage").textContent = "Nog geen triage voor deze mail";
    showTriageState(null);
    if (!t) return;
    $("triageAt").textContent = formatAt(t.at);
    // Pyramid layout: headline first, then short sections with bullets. **text** renders bold.
    $("triageHeadline").innerHTML = rich(t.headline || t.summary || "");
    var body = $("triageBody");
    body.innerHTML = "";
    var secs = [["Situatie", t.situation], ["Wat speelt", t.complication]];
    if (!t.situation && !t.complication) secs = [["Context", t.context ? [t.context] : null]];
    secs.forEach(function (s) {
      var items = s[1];
      if (!items || !items.length) return;
      if (!Array.isArray(items)) items = [items];
      var d = document.createElement("div");
      d.className = "tsec";
      d.innerHTML = '<h3 class="tsec-h">' + s[0] + '</h3><ul>' + items.map(function (x) { return "<li>" + rich(x) + "</li>"; }).join("") + "</ul>";
      body.appendChild(d);
    });
    // Kind tag: FYI / actie voor iemand anders / actie voor jou, readable at a glance.
    var KINDS = { fyi: "FYI", other: "Actie voor ander", me: "Actie voor jou" };
    var kt = $("triageKind");
    kt.hidden = !KINDS[t.kind];
    kt.textContent = KINDS[t.kind] || "";
    kt.setAttribute("data-kind", t.kind || "");
    // Notion lanes live below the card, so the card itself never offers a Notion action.
    var notNotion = function (a) { return a && a.type !== "notion"; };
    var rec = notNotion(t.recommend) ? t.recommend : ((t.alternatives || []).filter(notNotion)[0] || { type: "none" });
    t = Object.assign({}, t, { recommend: rec, alternatives: (t.alternatives || []).filter(function (a) { return notNotion(a) && a !== rec; }) });
    currentTriage = t;
    $("triagePrimary").textContent = actionLabel(rec);
    $("triageReason").hidden = !rec.reason;
    $("triageReason").innerHTML = rich(rec.reason || "");
    // Reply options (A kort, B warm, C afhouden): each is its own button with the gist visible.
    var opts = $("triageOptions");
    opts.innerHTML = "";
    var hasOpts = rec.type === "draft" && rec.options && rec.options.length;
    opts.hidden = !hasOpts;
    $("triagePrimary").hidden = !!hasOpts;
    if (hasOpts) rec.options.forEach(function (o) {
      var b = document.createElement("button");
      b.type = "button";
      b.className = "opt";
      b.innerHTML = "<b>" + rich(o.key + ". " + o.label) + "</b><span>" + rich(o.gist) + "</span>";
      b.addEventListener("click", function () { runAction({ type: "draft", option: o.key, gist: o.gist, label: "Draft " + o.key + " (" + o.label + ")" }); });
      opts.appendChild(b);
    });
    // Own direction: Leon types what the reply should say, EA writes the draft.
    if (rec.type === "draft") {
      var wrap = document.createElement("div");
      wrap.className = "opt-own";
      wrap.innerHTML = '<textarea id="ownGist" rows="2" placeholder="Eigen richting, bijv.: klinkt interessant, vraag naar fees en talen"></textarea>' +
        '<button type="button" class="btn-alt" id="ownGo">Draft met mijn richting</button>';
      opts.appendChild(wrap);
      opts.hidden = false;
      $("triagePrimary").hidden = true;
      wrap.querySelector("#ownGo").addEventListener("click", function () {
        var g = wrap.querySelector("#ownGist").value.trim();
        if (!g) { setStatus("Typ eerst kort wat het antwoord moet zeggen.", true); return; }
        runAction({ type: "draft", option: "eigen", gist: g, label: "Draft met eigen richting" });
      });
    }
    var alts = $("triageAlts");
    alts.innerHTML = "";
    var list = (t.alternatives || []).slice();
    if (rec.type !== "none" && !list.some(function (x) { return x.type === "none"; })) list.push({ type: "none" });
    list.forEach(function (a) {
      var b = document.createElement("button");
      b.type = "button";
      b.className = "btn-alt";
      b.textContent = actionLabel(a);
      if (a.reason) b.title = a.reason;
      b.addEventListener("click", function () { runAction(a); });
      alts.appendChild(b);
    });
    if (action) showTriageState(pendingText(action), true, clearRecordedAction);
  }

  // ---- UI ------------------------------------------------------------------
  function renderLanes() {
    var wrap = $("lanes");
    wrap.innerHTML = "";
    LANES.forEach(function (lane) {
      var b = document.createElement("button");
      b.type = "button";
      b.className = "lane";
      b.setAttribute("data-key", lane.key);
      b.setAttribute("data-tone", lane.tone);
      b.setAttribute("aria-pressed", "false");
      b.title = lane.category;
      b.innerHTML = '<span class="pill" data-tone="' + lane.tone + '">' + lane.label + '</span>';
      b.addEventListener("click", function () { pick(lane); });
      wrap.appendChild(b);
    });
  }

  function showDone(lane, note) {
    var buttons = document.querySelectorAll(".lane");
    for (var i = 0; i < buttons.length; i++) {
      var on = !!lane && buttons[i].getAttribute("data-key") === lane.key;
      buttons[i].classList.toggle("active", on);
      buttons[i].setAttribute("aria-pressed", on ? "true" : "false");
    }
    $("done").hidden = !lane;
    if (!lane) return;
    $("doneText").textContent = "Staat klaar voor EA: " + lane.label;
    var n = $("doneNote");
    n.hidden = !note;
    n.textContent = note ? "Notitie: " + note : "";
  }

  function setBusy(on) {
    busy = on;
    var buttons = document.querySelectorAll(".lane, .undo, .btn-primary, .btn-alt");
    for (var i = 0; i < buttons.length; i++) buttons[i].disabled = on;
  }

  function setStatus(text, isError) {
    var s = $("status");
    s.textContent = text || "";
    s.classList.toggle("error", !!isError);
  }

  function loadSubject() {
    var it = item();
    if (!it) { $("subject").textContent = "Geen mail geselecteerd"; return; }
    if (typeof it.subject === "string") { $("subject").textContent = it.subject || "(geen onderwerp)"; return; }
    it.subject.getAsync(function (r) {
      $("subject").textContent = (r.status === Office.AsyncResultStatus.Succeeded && r.value) ? r.value : "Nieuwe mail";
    });
  }

  // Show the existing state when the pane opens or the selected mail changes.
  var refreshSeq = 0;
  function refresh(interactive) {
    var seq = ++refreshSeq;
    setStatus("");
    $("note").value = "";
    showDone(null);
    currentTriage = null;
    $("triage").hidden = true;
    $("noTriage").hidden = true;
    $("signin").hidden = true;
    loadSubject();
    if (!item()) return;
    if (!readMode()) {
      // Compose: lanes and note stay local.
      Promise.all([localCategories().catch(function () { return []; }), localProps().catch(function () { return null; })])
        .then(function (res) {
          if (seq !== refreshSeq) return;
          var lanes = res[0].filter(function (n) { return !!laneByCategory(n); });
          var note = res[1] ? (res[1].get(NOTE_KEY) || "") : "";
          $("note").value = note;
          showDone(lanes.length ? laneByCategory(lanes[0]) : null, note);
        });
      return;
    }
    $("noTriage").hidden = false;
    $("noTriage").textContent = "Triage ophalen...";
    var id;
    try { id = restId(); } catch (e) { $("noTriage").textContent = "Triage ophalen mislukt: " + msg(e); return; }
    getMessage(id, interactive === true ? true : false).then(function (m) {
      if (seq !== refreshSeq) return; // selection moved on
      var lanes = m.categories.filter(function (n) { return !!laneByCategory(n); });
      var note = m.props[NOTE_KEY] || "";
      $("note").value = note;
      showDone(lanes.length ? laneByCategory(lanes[0]) : null, note);
      renderTriage(parseJson(m.props[TRIAGE_KEY]), parseJson(m.props[ACTION_KEY]));
    }).catch(function (e) {
      if (seq !== refreshSeq) return;
      renderTriage(null, null);
      if (e && e.code === "LOGIN_REQUIRED") {
        $("noTriage").textContent = "Meld je eenmalig aan om de triage te zien.";
        $("signin").hidden = false;
        return;
      }
      $("noTriage").textContent = e && e.code === "NAA_UNSUPPORTED" ? e.message : "Triage ophalen mislukt: " + msg(e);
    });
  }

  Office.onReady(function (info) {
    renderLanes();
    $("undo").addEventListener("click", undo);
    $("signin").addEventListener("click", function () { refresh(true); });
    $("triageUndo").addEventListener("click", undoTriageState);
    $("recentUndo").addEventListener("click", undoRecent);
    $("triagePrimary").addEventListener("click", function () {
      if (currentTriage) runAction(currentTriage.recommend || { type: "none" });
    });
    // The note sits below the lanes, so it is also saved when typed after picking a lane.
    var noteTimer = null;
    $("note").addEventListener("input", function () {
      clearTimeout(noteTimer);
      noteTimer = setTimeout(function () {
        var active = document.querySelector(".lane.active");
        if (!active || busy) return;
        var lane = LANES.filter(function (l) { return l.key === active.getAttribute("data-key"); })[0];
        var note = $("note").value.trim();
        saveNote(note, lane).then(function () {
          showDone(lane, note);
          setStatus("Notitie bewaard.");
        }, function (err) { setStatus("De notitie kon niet worden opgeslagen: " + msg(err), true); });
      }, 800);
    });
    if (info.host !== Office.HostType.Outlook) {
      setStatus("Open dit paneel vanuit Outlook.", true);
      setBusy(true);
      return;
    }
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
      setStatus("Deze Outlook-versie ondersteunt geen categorieën via add-ins (Mailbox 1.8 nodig).", true);
      setBusy(true);
      return;
    }
    if (!naaSupported()) {
      setStatus("Deze Outlook-versie ondersteunt geen aanmelden via de add-in (Nested App Authentication). Triage en archiveren werken hier niet.", true);
    }
    refresh();
    // Pinned pane: follow the selection to the next mail.
    if (Office.EventType && Office.EventType.ItemChanged) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { refresh(); });
    }
  });
})();
