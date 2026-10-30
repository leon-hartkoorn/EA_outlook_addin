/* Eva: Naar Notion. Shows Eva's triage for the selected mail (custom property
   evaTriage) with one-click actions, and tags the mail with a "Notion: <lane>"
   category (and optional note) so Eva can turn it into a Notion card later.
   No tokens, no network calls besides Office.js (incl. EWS via the mailbox). */
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

  function getItemCategories() {
    return call(function (cb) { item().categories.getAsync(cb); }).then(function (v) { return v || []; });
  }
  function getMasterCategories() {
    return call(function (cb) { Office.context.mailbox.masterCategories.getAsync(cb); }).then(function (v) { return v || []; });
  }
  function loadProps() {
    return call(function (cb) { item().loadCustomPropertiesAsync(cb); });
  }

  function laneByCategory(name) {
    for (var i = 0; i < LANES.length; i++) if (LANES[i].category === name) return LANES[i];
    return null;
  }

  // ---- Core actions --------------------------------------------------------
  function ensureMasterCategories(list) {
    list = list || LANES.map(function (l) { return { displayName: l.category, color: l.color }; });
    return getMasterCategories().then(function (existing) {
      var have = {};
      existing.forEach(function (c) { have[c.displayName.toLowerCase()] = true; });
      var missing = list.filter(function (m) { return !have[m.displayName.toLowerCase()]; });
      if (!missing.length) return;
      return call(function (cb) { Office.context.mailbox.masterCategories.addAsync(missing, cb); })
        .catch(function (err) {
          // Another client may have added one in the meantime; re-check before failing.
          return getMasterCategories().then(function (now) {
            var names = now.map(function (c) { return c.displayName.toLowerCase(); });
            var stillMissing = missing.filter(function (m) { return names.indexOf(m.displayName.toLowerCase()) < 0; });
            if (stillMissing.length) throw err;
          });
        });
    });
  }

  function currentNotionCategories() {
    return getItemCategories().then(function (cats) {
      return cats.map(function (c) { return c.displayName; }).filter(function (n) { return !!laneByCategory(n); });
    });
  }

  function removeCategories(names) {
    if (!names.length) return Promise.resolve();
    return call(function (cb) { item().categories.removeAsync(names, cb); });
  }

  // Note is best effort: the category alone is enough for Eva.
  function saveNote(text, lane) {
    return loadProps().then(function (props) {
      if (text) props.set(NOTE_KEY, text); else props.remove(NOTE_KEY);
      if (lane) props.set(LANE_KEY, lane.category); else props.remove(LANE_KEY);
      return call(function (cb) { props.saveAsync(cb); });
    }).then(function () { return true; }, function () { return false; });
  }

  function pick(lane) {
    if (busy) return;
    var note = $("note").value.trim();
    setBusy(true);
    setStatus("Bezig...");
    ensureMasterCategories()
      .then(currentNotionCategories)
      .then(function (present) {
        var others = present.filter(function (n) { return n !== lane.category; });
        return removeCategories(others).then(function () {
          if (present.indexOf(lane.category) >= 0) return;
          return call(function (cb) { item().categories.addAsync([lane.category], cb); });
        });
      })
      .then(function () { return saveNote(note, lane); })
      .then(function (noteSaved) {
        showDone(lane, note);
        if (note && !noteSaved) setStatus("Lane staat erop. De notitie kon niet worden opgeslagen.", true);
        else setStatus("");
      })
      .catch(function (err) { setStatus("Dat lukte niet: " + msg(err), true); })
      .then(function () { setBusy(false); });
  }

  function undo() {
    if (busy) return;
    setBusy(true);
    setStatus("Ongedaan maken...");
    currentNotionCategories()
      .then(removeCategories)
      .then(function () { return saveNote("", null); })
      .then(function () {
        showDone(null);
        $("note").value = "";
        setStatus("Verwijderd. Eva pakt deze mail niet op.");
      })
      .catch(function (err) { setStatus("Ongedaan maken lukte niet: " + msg(err), true); })
      .then(function () { setBusy(false); });
  }

  // ---- Eva triage ----------------------------------------------------------
  // Eva writes "evaTriage" (a JSON string) into the custom properties of a mail.
  // A click is recorded as "evaAction" plus the category "Eva: actie"; Eva's
  // poller executes it. Archive and "gelezen" are also tried directly via EWS.
  var TRIAGE_KEY = "evaTriage";
  var ACTION_KEY = "evaAction";
  var ACTION_CAT = { displayName: "Eva: actie", color: "Preset8" };
  var M_NS = "http://schemas.microsoft.com/exchange/services/2006/messages";
  var T_NS = "http://schemas.microsoft.com/exchange/services/2006/types";
  var folderCache = {};
  var recentUndo = null;   // undo for a real EWS move; survives a mail switch
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
      case "draft": return "Laat Eva een reactie schrijven";
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

  // ---- EWS (classic Outlook, Outlook on the web) ---------------------------
  function xmlEsc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function ewsAvailable() {
    return !!(Office.context.mailbox && typeof Office.context.mailbox.makeEwsRequestAsync === "function" && item() && item().itemId);
  }
  function ews(body) {
    var env = '<?xml version="1.0" encoding="utf-8"?>' +
      '<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="' + M_NS + '" xmlns:t="' + T_NS + '">' +
      '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
      '<soap:Body>' + body + '</soap:Body></soap:Envelope>';
    return new Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.makeEwsRequestAsync(env, function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) { reject(r.error || new Error("EWS mislukt")); return; }
          var doc = new DOMParser().parseFromString(r.value, "text/xml");
          var all = doc.getElementsByTagName("*");
          for (var i = 0; i < all.length; i++) {
            var rc = all[i].getAttribute && all[i].getAttribute("ResponseClass");
            if (rc && rc !== "Success") {
              var t = all[i].getElementsByTagNameNS(M_NS, "MessageText")[0];
              reject(new Error(t ? t.textContent : "EWS fout"));
              return;
            }
          }
          resolve(doc);
        });
      } catch (e) { reject(e); }
    });
  }
  function findFolderId(path) {
    if (folderCache[path]) return Promise.resolve(folderCache[path]);
    var parts = path.split("/").map(function (p) { return p.trim(); }).filter(Boolean);
    var parent = '<t:DistinguishedFolderId Id="msgfolderroot"/>';
    var chain = Promise.resolve(null);
    parts.forEach(function (name) {
      chain = chain.then(function () {
        return ews('<m:FindFolder Traversal="Shallow"><m:FolderShape><t:BaseShape>IdOnly</t:BaseShape></m:FolderShape>' +
          '<m:Restriction><t:IsEqualTo><t:FieldURI FieldURI="folder:DisplayName"/><t:FieldURIOrConstant>' +
          '<t:Constant Value="' + xmlEsc(name) + '"/></t:FieldURIOrConstant></t:IsEqualTo></m:Restriction>' +
          '<m:ParentFolderIds>' + parent + '</m:ParentFolderIds></m:FindFolder>').then(function (doc) {
          var f = doc.getElementsByTagNameNS(T_NS, "FolderId")[0];
          if (!f) throw new Error("Map niet gevonden: " + name);
          parent = '<t:FolderId Id="' + xmlEsc(f.getAttribute("Id")) + '"/>';
          return f.getAttribute("Id");
        });
      });
    });
    return chain.then(function (id) { folderCache[path] = id; return id; });
  }
  function ewsIsRead(id) {
    return ews('<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>' +
      '<t:FieldURI FieldURI="message:IsRead"/></t:AdditionalProperties></m:ItemShape>' +
      '<m:ItemIds><t:ItemId Id="' + xmlEsc(id) + '"/></m:ItemIds></m:GetItem>').then(function (doc) {
      var n = doc.getElementsByTagNameNS(T_NS, "IsRead")[0];
      return n ? n.textContent === "true" : false;
    });
  }
  function ewsSetRead(id, read) {
    return ews('<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite"><m:ItemChanges><t:ItemChange>' +
      '<t:ItemId Id="' + xmlEsc(id) + '"/><t:Updates><t:SetItemField><t:FieldURI FieldURI="message:IsRead"/>' +
      '<t:Message><t:IsRead>' + (read ? "true" : "false") + '</t:IsRead></t:Message></t:SetItemField></t:Updates>' +
      '</t:ItemChange></m:ItemChanges></m:UpdateItem>');
  }
  function ewsMove(id, folderXml) {
    return ews('<m:MoveItem><m:ToFolderId>' + folderXml + '</m:ToFolderId>' +
      '<m:ItemIds><t:ItemId Id="' + xmlEsc(id) + '"/></m:ItemIds></m:MoveItem>').then(function (doc) {
      var n = doc.getElementsByTagNameNS(T_NS, "ItemId")[0];
      return n ? n.getAttribute("Id") : null;
    });
  }

  // ---- Recorded actions (all Outlook versions) -----------------------------
  function recordAction(a) {
    var payload = { type: a.type, params: actionParams(a), label: actionLabel(a), at: new Date().toISOString() };
    return ensureMasterCategories([ACTION_CAT])
      .then(loadProps)
      .then(function (props) {
        props.set(ACTION_KEY, JSON.stringify(payload));
        return call(function (cb) { props.saveAsync(cb); });
      })
      .then(getItemCategories)
      .then(function (cats) {
        var has = cats.some(function (c) { return c.displayName === ACTION_CAT.displayName; });
        if (!has) return call(function (cb) { item().categories.addAsync([ACTION_CAT.displayName], cb); });
      })
      .then(function () { return payload; });
  }
  function clearRecordedAction() {
    return getItemCategories()
      .then(function (cats) {
        var has = cats.some(function (c) { return c.displayName === ACTION_CAT.displayName; });
        if (has) return call(function (cb) { item().categories.removeAsync([ACTION_CAT.displayName], cb); });
      })
      .then(loadProps)
      .then(function (props) {
        props.remove(ACTION_KEY);
        return call(function (cb) { props.saveAsync(cb); });
      });
  }

  function pendingText(a) {
    switch (a.type) {
      case "archive": return "Eva verplaatst hem zo naar " + ((a.params && (a.params.folderLabel || a.params.folderPath)) || "de map");
      case "draft": return "Staat klaar voor Eva: zij schrijft een concept";
      case "fyi": return "Staat klaar voor Eva: wordt als gelezen gemarkeerd";
      default: return "Staat klaar voor Eva: " + (a.label || a.type);
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
    var itemId = item().itemId;
    setBusy(true);
    setStatus("Bezig...");
    var direct = Promise.reject(new Error("geen EWS"));
    if (ewsAvailable() && (a.type === "archive" || a.type === "fyi")) {
      direct = ewsIsRead(itemId).then(function (wasRead) {
        if (a.type === "fyi") {
          return ewsSetRead(itemId, true).then(function () {
            return { text: "Gemarkeerd als gelezen", undo: function () { return ewsSetRead(itemId, wasRead); } };
          });
        }
        if (!a.folderPath) throw new Error("geen map");
        // Resolve the folder first, so a missing folder falls back before anything changes.
        return findFolderId(a.folderPath).then(function (fid) {
          return ewsSetRead(itemId, true).then(function () {
            return ewsMove(itemId, '<t:FolderId Id="' + xmlEsc(fid) + '"/>');
          });
        }).then(function (newId) {
          return {
            moved: true,
            text: "Gearchiveerd in " + (a.folderLabel || a.folderPath),
            undo: function () {
              if (!newId) return Promise.reject(new Error("terugzetten kan niet, zoek hem in " + a.folderPath));
              return ewsMove(newId, '<t:DistinguishedFolderId Id="inbox"/>').then(function (backId) {
                if (!wasRead && backId) return ewsSetRead(backId, false);
              });
            }
          };
        });
      });
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
      }, function () {
        // Fallback: record the action for Eva's poller.
        return recordAction(a).then(function (payload) {
          showTriageState(pendingText(payload), true, clearRecordedAction);
          setStatus("");
        });
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

  function renderTriage(t, action) {
    currentTriage = t;
    var readMode = !!(item() && item().itemId); // triage only applies to a received mail
    $("triage").hidden = !t || !readMode;
    $("noTriage").hidden = !!t || !readMode;
    showTriageState(null);
    if (!t) return;
    $("triageAt").textContent = formatAt(t.at);
    $("triageSummary").textContent = t.summary || "";
    $("triageContext").hidden = !t.context;
    $("triageContext").textContent = t.context || "";
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
    $("triageReason").textContent = rec.reason || "";
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
    $("doneText").textContent = "Staat klaar voor Eva: " + lane.label;
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

  function msg(err) { return (err && (err.message || err.name)) || "onbekende fout"; }

  function loadSubject() {
    var it = item();
    if (!it) { $("subject").textContent = "Geen mail geselecteerd"; return; }
    if (typeof it.subject === "string") { $("subject").textContent = it.subject || "(geen onderwerp)"; return; }
    it.subject.getAsync(function (r) {
      $("subject").textContent = (r.status === Office.AsyncResultStatus.Succeeded && r.value) ? r.value : "Nieuwe mail";
    });
  }

  // Show the existing state when the pane opens or the selected mail changes.
  function refresh() {
    setStatus("");
    $("note").value = "";
    showDone(null);
    currentTriage = null;
    $("triage").hidden = true;
    $("noTriage").hidden = true;
    if (!item()) { loadSubject(); return; }
    loadSubject();
    Promise.all([
      currentNotionCategories(),
      loadProps().catch(function () { return null; })
    ]).then(function (res) {
      var p = res[1];
      var note = p ? (p.get(NOTE_KEY) || "") : "";
      var lane = res[0].length ? laneByCategory(res[0][0]) : null;
      $("note").value = note;
      showDone(lane, note);
      renderTriage(p ? parseJson(p.get(TRIAGE_KEY)) : null, p ? parseJson(p.get(ACTION_KEY)) : null);
    }).catch(function () { renderTriage(null, null); });
  }

  Office.onReady(function (info) {
    renderLanes();
    $("undo").addEventListener("click", undo);
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
        saveNote(note, lane).then(function (ok) {
          if (ok) { showDone(lane, note); setStatus("Notitie bewaard."); } else setStatus("De notitie kon niet worden opgeslagen.", true);
        });
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
    refresh();
    // Pinned pane: follow the selection to the next mail.
    if (Office.EventType && Office.EventType.ItemChanged) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, refresh);
    }
  });
})();
