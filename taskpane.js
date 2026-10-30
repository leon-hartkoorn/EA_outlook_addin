/* Eva: Naar Notion. Tags the current mail with a "Notion: <lane>" category
   (and optional note) so Eva can turn it into a Notion card later.
   No tokens, no network calls besides Office.js itself. */
(function () {
  "use strict";

  var LANES = [
    // Notion names and colours; order by how often Leon uses them.
    { key: "today",   label: "Sprint 1D", sub: "vandaag",               category: "Notion: Sprint 1 day",        color: "Preset0",  tone: "red" },
    { key: "nextday", label: "Next day",  sub: "morgen",                category: "Notion: Sprint next day",     color: "Preset7",  tone: "blue" },
    { key: "week",    label: "Sprint 1W", sub: "deze week",             category: "Notion: Sprint 1 week",       color: "Preset7",  tone: "blue" },
    { key: "backlog", label: "Backlog",   sub: "ooit, nog geen sprint", category: "Notion: Backlog",             color: "Preset12", tone: "default" },
    { key: "onhold",  label: "On hold",   sub: "wacht op iemand anders", category: "Notion: On hold/waiting for", color: "Preset7",  tone: "blue" }
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
  function ensureMasterCategories() {
    return getMasterCategories().then(function (existing) {
      var have = {};
      existing.forEach(function (c) { have[c.displayName.toLowerCase()] = true; });
      var missing = LANES.filter(function (l) { return !have[l.category.toLowerCase()]; })
        .map(function (l) { return { displayName: l.category, color: l.color }; });
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
      b.innerHTML = '<span class="pill" data-tone="' + lane.tone + '">' + lane.label + '</span>' +
                    '<span class="sub">' + lane.sub + '</span>';
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
    var buttons = document.querySelectorAll(".lane, .undo");
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
    if (!item()) { loadSubject(); return; }
    loadSubject();
    Promise.all([
      currentNotionCategories(),
      loadProps().then(function (p) { return p.get(NOTE_KEY) || ""; }, function () { return ""; })
    ]).then(function (res) {
      var lane = res[0].length ? laneByCategory(res[0][0]) : null;
      $("note").value = res[1];
      showDone(lane, res[1]);
    }).catch(function () { /* read failures just leave the empty state */ });
  }

  Office.onReady(function (info) {
    renderLanes();
    $("undo").addEventListener("click", undo);
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
