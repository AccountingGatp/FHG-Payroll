/* app.js — browser UI glue for the Payroll → QuickBooks converter.
 * All parsing and journal-building lives in converter.js; this file only
 * handles file input, the confirm step, and rendering / download. */
(function () {
  "use strict";
  var PC = window.PayrollConverter;
  var XLSX = window.XLSX;

  var state = {
    input: null, // { headerDate, rows }
    master: null, // { books, bookList, memoBooks }
    detection: null,
    lastResult: null,
    downloadBlobUrl: null,
    downloadName: null,
  };

  var $ = function (id) {
    return document.getElementById(id);
  };

  // ---- File loading --------------------------------------------------------

  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () {
        resolve(fr.result);
      };
      fr.onerror = function () {
        reject(new Error("Could not read " + file.name));
      };
      fr.readAsArrayBuffer(file);
    });
  }

  function wireDropzone(zoneId, inputId, statusId, onFile) {
    var zone = $(zoneId);
    var input = $(inputId);
    var status = $(statusId);

    input.addEventListener("change", function () {
      if (input.files && input.files[0]) handle(input.files[0]);
    });
    ["dragenter", "dragover"].forEach(function (ev) {
      zone.addEventListener(ev, function (e) {
        e.preventDefault();
        zone.classList.add("dragover");
      });
    });
    ["dragleave", "drop"].forEach(function (ev) {
      zone.addEventListener(ev, function (e) {
        e.preventDefault();
        zone.classList.remove("dragover");
      });
    });
    zone.addEventListener("drop", function (e) {
      var f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handle(f);
    });

    function handle(file) {
      status.textContent = "Reading " + file.name + " …";
      readFile(file)
        .then(function (buf) {
          onFile(buf, file.name);
          zone.classList.add("loaded");
          status.textContent = "✓ " + file.name;
        })
        .catch(function (err) {
          zone.classList.remove("loaded");
          status.textContent = "Click or drop file";
          showFileError(err.message);
        });
    }
  }

  function showFileError(msg) {
    var el = $("file-error");
    el.textContent = msg;
    el.hidden = false;
  }
  function clearFileError() {
    $("file-error").hidden = true;
  }

  // ---- Wiring --------------------------------------------------------------

  wireDropzone("dz-input", "file-input", "status-input", function (buf, name) {
    clearFileError();
    try {
      state.input = PC.parseInput(XLSX, buf);
    } catch (e) {
      state.input = null;
      $("dz-input").classList.remove("loaded");
      $("status-input").textContent = "Click or drop file";
      showFileError("Payroll input: " + e.message);
      return;
    }
    maybeShowConfirm();
  });

  wireDropzone("dz-master", "file-master", "status-master", function (buf, name) {
    clearFileError();
    try {
      state.master = PC.parseMaster(XLSX, buf);
    } catch (e) {
      state.master = null;
      $("dz-master").classList.remove("loaded");
      $("status-master").textContent = "Click or drop file";
      showFileError("Master Account List: " + e.message);
      return;
    }
    maybeShowConfirm();
  });

  function maybeShowConfirm() {
    if (!state.input || !state.master) return;

    // Populate book select + recommendation
    state.detection = PC.detectBook(state.input.rows, state.master);
    var sel = $("book-select");
    sel.innerHTML = "";
    state.master.bookList.forEach(function (b) {
      var o = document.createElement("option");
      o.value = b;
      o.textContent =
        b + " (" + state.detection.coverage[b] + "/" + state.detection.totalMemos + " memos map)";
      if (b === state.detection.recommended) o.selected = true;
      sel.appendChild(o);
    });
    $("book-reason").textContent = "Recommendation: " + state.detection.reason;

    // Dates
    var dates = PC.computeDates(state.input.headerDate);
    $("new-date").value = toInputDate(dates.newDate);
    $("old-date").value = toInputDate(dates.oldDate);
    updateJournalNos();

    // Class IDs
    var ids = PC.collectClassIds(state.input.rows);
    $("class-count").textContent = ids.length + " found";
    var box = $("detected-ids");
    box.innerHTML = "";
    ids.forEach(function (id) {
      var chip = document.createElement("span");
      chip.className = "id-chip";
      chip.textContent = id;
      box.appendChild(chip);
    });

    $("step-confirm").hidden = false;
    refreshUnmappedPreview();
    $("step-confirm").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function updateJournalNos() {
    var nd = fromInputDate($("new-date").value);
    var od = fromInputDate($("old-date").value);
    $("new-jno").textContent = nd ? "Journal No: " + PC.journalNo(nd) : "";
    $("old-jno").textContent = od ? "Journal No: " + PC.journalNo(od) : "";
  }

  function refreshUnmappedPreview() {
    var book = $("book-select").value;
    var accountMap = (state.master.books && state.master.books[book]) || {};
    var seen = {};
    var unmapped = [];
    state.input.rows.forEach(function (r) {
      if (seen[r.memo]) return;
      seen[r.memo] = true;
      if (accountMap[r.memo] === undefined) unmapped.push(r.memo);
    });
    var el = $("unmapped-preview");
    if (unmapped.length) {
      el.innerHTML =
        "<strong>" +
        unmapped.length +
        ' memo(s) have no account in the "' +
        book +
        '" book.</strong> They will be left blank and highlighted yellow in the workbook — resolve before importing:<ul>' +
        unmapped
          .map(function (m) {
            return "<li>" + escapeHtml(m) + "</li>";
          })
          .join("") +
        "</ul>";
      el.hidden = false;
    } else {
      el.hidden = true;
    }
  }

  $("book-select").addEventListener("change", refreshUnmappedPreview);
  $("new-date").addEventListener("change", updateJournalNos);
  $("old-date").addEventListener("change", updateJournalNos);

  // ---- Class map parsing ---------------------------------------------------

  function parseClassMap() {
    var text = $("class-map").value || "";
    var map = {};
    text.split(/\r?\n/).forEach(function (line) {
      var t = line.trim();
      if (!t) return;
      var m = t.match(/^([0-9]+)\s*[=,:\t]\s*(.+)$/);
      if (m) map[m[1].trim()] = m[2].trim();
    });
    return Object.keys(map).length ? map : null;
  }

  // ---- Generate ------------------------------------------------------------

  $("btn-generate").addEventListener("click", function () {
    var book = $("book-select").value;
    var nd = fromInputDate($("new-date").value);
    var od = fromInputDate($("old-date").value);
    if (!nd || !od) {
      showFileError("Please set both journal dates.");
      return;
    }
    var classMap = parseClassMap();

    // Build with explicit dates from the (possibly edited) inputs.
    var newSheet = PC.buildSheet(state.input.rows, book, state.master, "new", nd, classMap);
    var oldSheet = PC.buildSheet(state.input.rows, book, state.master, "old", od, classMap);
    var meta = { book: book, generated: nowStamp() };
    var wb = PC.generateWorkbook(XLSX, newSheet, oldSheet, meta);

    var out = XLSX.write(wb, { type: "array", bookType: "xlsx", cellDates: true });
    var blob = new Blob([out], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    if (state.downloadBlobUrl) URL.revokeObjectURL(state.downloadBlobUrl);
    state.downloadBlobUrl = URL.createObjectURL(blob);
    state.downloadName =
      "Payroll_QB_Import_" +
      book +
      "_" +
      PC.mmddyyyy(nd) +
      ".xlsx";

    state.lastResult = { newSheet: newSheet, oldSheet: oldSheet, book: book };
    renderResult();

    // Auto-download for convenience.
    triggerDownload();
  });

  function triggerDownload() {
    var a = document.createElement("a");
    a.href = state.downloadBlobUrl;
    a.download = state.downloadName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  $("btn-download").addEventListener("click", triggerDownload);

  // ---- Rendering result ----------------------------------------------------

  function renderResult() {
    var r = state.lastResult;
    var cards = $("summary-cards");
    cards.innerHTML = "";
    [
      { label: "New Month", s: r.newSheet },
      { label: "Old Month", s: r.oldSheet },
    ].forEach(function (item) {
      cards.appendChild(sumCard(item.label, item.s));
    });

    // Combined unmapped list
    var un = {};
    r.newSheet.summary.unmapped.concat(r.oldSheet.summary.unmapped).forEach(function (m) {
      un[m] = true;
    });
    var list = Object.keys(un);
    var el = $("result-unmapped");
    if (list.length) {
      el.innerHTML =
        "<strong>Unmapped memos (highlighted yellow in the workbook — resolve before import):</strong><ul>" +
        list
          .map(function (m) {
            return "<li>" + escapeHtml(m) + "</li>";
          })
          .join("") +
        "</ul>";
      el.hidden = false;
    } else {
      el.hidden = true;
    }

    $("download-name").textContent = state.downloadName;
    $("step-result").hidden = false;
    $("step-result").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function sumCard(label, s) {
    var sum = s.summary;
    var balanced = Math.abs(sum.difference) < 0.005;
    var div = document.createElement("div");
    div.className = "sum-card";
    div.innerHTML =
      "<h3>" +
      label +
      "</h3>" +
      '<div class="jno">' +
      s.journalNo +
      " · " +
      fmtDate(s.journalDate) +
      "</div>" +
      '<div class="badge ' +
      (balanced ? "ok" : "bad") +
      '">' +
      (balanced ? "BALANCED" : "OUT OF BALANCE") +
      "</div>" +
      '<table class="sum-rows">' +
      row("Lines", sum.lineCount, false) +
      row("Total Debit", money(sum.totalDebit), true) +
      row("Total Credit", money(sum.totalCredit), true) +
      rowTotal("Difference", money(sum.difference)) +
      "</table>" +
      '<p class="hint">Clearing: ' +
      money(sum.clearingBefore) +
      " → " +
      money(sum.clearingAfter) +
      " (adj " +
      money(sum.adjustment) +
      ")</p>";
    return div;
  }

  function row(label, val, num) {
    return (
      "<tr><td>" +
      label +
      '</td><td class="' +
      (num ? "num" : "num") +
      '">' +
      val +
      "</td></tr>"
    );
  }
  function rowTotal(label, val) {
    return (
      '<tr class="total"><td>' + label + '</td><td class="num">' + val + "</td></tr>"
    );
  }

  // ---- Helpers -------------------------------------------------------------

  function toInputDate(d) {
    return (
      d.getFullYear() +
      "-" +
      String(d.getMonth() + 1).padStart(2, "0") +
      "-" +
      String(d.getDate()).padStart(2, "0")
    );
  }
  function fromInputDate(s) {
    if (!s) return null;
    var p = s.split("-");
    if (p.length !== 3) return null;
    return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  }
  function fmtDate(d) {
    return (
      String(d.getMonth() + 1).padStart(2, "0") +
      "/" +
      String(d.getDate()).padStart(2, "0") +
      "/" +
      d.getFullYear()
    );
  }
  function money(n) {
    var neg = n < 0;
    var v = Math.abs(n).toLocaleString("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
    return neg ? "(" + v + ")" : v;
  }
  function nowStamp() {
    var d = new Date();
    return (
      d.getFullYear() +
      "-" +
      String(d.getMonth() + 1).padStart(2, "0") +
      "-" +
      String(d.getDate()).padStart(2, "0")
    );
  }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[c];
    });
  }
})();
