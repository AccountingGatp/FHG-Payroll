/*
 * converter.js — Payroll → QuickBooks Journal Import core logic
 *
 * Pure, framework-agnostic module. Runs in the browser (attaches to
 * window.PayrollConverter) and in Node (module.exports) so the exact same
 * code path is unit-tested and shipped. The SheetJS-compatible library
 * (xlsx-js-style) is passed in by the caller — this module never imports it.
 *
 * Implements the GATP "Payroll → QuickBooks Journal Import" SOP.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.PayrollConverter = api;
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var CLEARING_MEMO = "Payroll Clearing Account";
  var CURRENCY_FMT = '#,##0.00;(#,##0.00);"-"';
  var DATE_FMT = "mm/dd/yyyy";
  var HEADER_FILL = "1F3864"; // GATP navy
  var ALT_FILL = "F2F2F2";
  var UNMAPPED_FILL = "FFFF00"; // yellow

  function round2(n) {
    return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
  }

  function isBlank(v) {
    return v === null || v === undefined || String(v).trim() === "";
  }

  // ---- Parsing -------------------------------------------------------------

  // Read the payroll input workbook (first sheet). Returns { headerDate, rows }.
  // headerDate is the YYYYMMDD value in cell A1. rows carry class/code/memo and
  // the New-month (col F) and Old-month (col G) amounts.
  function parseInput(XLSX, arrayBuffer) {
    var wb = XLSX.read(arrayBuffer, { type: "array" });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
    if (!aoa.length) throw new Error("Payroll input sheet is empty.");

    var headerDate = parseHeaderDate(aoa[0][0]);
    var rows = [];
    for (var i = 1; i < aoa.length; i++) {
      var r = aoa[i];
      var memo = r[3];
      if (isBlank(memo)) continue;
      rows.push({
        cls: r[1],
        code: r[2],
        memo: String(memo).trim(),
        orig: toNum(r[4]),
        newAmt: toNum(r[5]),
        oldAmt: toNum(r[6]),
      });
    }
    if (!headerDate)
      throw new Error(
        "Could not read a YYYYMMDD date from cell A1 of the payroll input."
      );
    return { headerDate: headerDate, rows: rows };
  }

  function toNum(v) {
    if (isBlank(v)) return 0;
    var n = Number(v);
    return isNaN(n) ? 0 : n;
  }

  // Accepts 20260501, "20260501", or an Excel/JS date; returns {y,m,d}.
  function parseHeaderDate(v) {
    if (v instanceof Date)
      return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
    var s = String(v).replace(/[^0-9]/g, "");
    if (s.length === 8) {
      return {
        y: parseInt(s.slice(0, 4), 10),
        m: parseInt(s.slice(4, 6), 10),
        d: parseInt(s.slice(6, 8), 10),
      };
    }
    return null;
  }

  // Read the Master Account List. Returns { books: {BookName: {memo:account}},
  // bookList: [...], memoBooks: {memo: [books]} }.
  function parseMaster(XLSX, arrayBuffer) {
    var wb = XLSX.read(arrayBuffer, { type: "array" });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
    var books = {};
    var bookList = [];
    var memoBooks = {};
    for (var i = 1; i < aoa.length; i++) {
      var book = aoa[i][0];
      var memo = aoa[i][1];
      var account = aoa[i][2];
      if (isBlank(book) || isBlank(memo)) continue;
      book = String(book).trim();
      memo = String(memo).trim();
      if (!books[book]) {
        books[book] = {};
        bookList.push(book);
      }
      books[book][memo] = account == null ? "" : String(account).trim();
      if (!memoBooks[memo]) memoBooks[memo] = [];
      if (memoBooks[memo].indexOf(book) === -1) memoBooks[memo].push(book);
    }
    return { books: books, bookList: bookList, memoBooks: memoBooks };
  }

  // ---- Book detection ------------------------------------------------------

  // Recommend which set of books to pull accounts from. Per the SOP: a memo that
  // exists in only one book points strongly at that book. Score each book by how
  // many input memos map ONLY to it, and by overall mapping coverage.
  function detectBook(inputRows, master) {
    var inputMemos = {};
    inputRows.forEach(function (r) {
      inputMemos[r.memo] = true;
    });
    var memos = Object.keys(inputMemos);

    var uniqueHits = {}; // book -> count of memos that exist ONLY in that book
    var coverage = {}; // book -> count of input memos it can map
    master.bookList.forEach(function (b) {
      uniqueHits[b] = 0;
      coverage[b] = 0;
    });

    var uniqueMemosByBook = {};
    memos.forEach(function (memo) {
      var owners = master.memoBooks[memo] || [];
      owners.forEach(function (b) {
        coverage[b]++;
      });
      if (owners.length === 1) {
        var b = owners[0];
        uniqueHits[b]++;
        (uniqueMemosByBook[b] = uniqueMemosByBook[b] || []).push(memo);
      }
    });

    // Pick the book with the most unique-memo hits; tie-break on coverage.
    var recommended = null;
    var best = -1;
    master.bookList.forEach(function (b) {
      var score = uniqueHits[b] * 1000 + coverage[b];
      if (score > best) {
        best = score;
        recommended = b;
      }
    });

    var reason;
    if (recommended && uniqueHits[recommended] > 0) {
      var ex = uniqueMemosByBook[recommended].slice(0, 3);
      reason =
        'Memo' +
        (ex.length > 1 ? "s" : "") +
        ' "' +
        ex.join('", "') +
        '" exist only in the ' +
        recommended +
        " book, so " +
        recommended +
        " is the likely answer.";
    } else if (recommended) {
      reason =
        recommended +
        " maps the most input memos (" +
        coverage[recommended] +
        " of " +
        memos.length +
        ").";
    } else {
      reason = "No book could be determined from the data.";
    }

    return {
      recommended: recommended,
      reason: reason,
      uniqueHits: uniqueHits,
      coverage: coverage,
      totalMemos: memos.length,
    };
  }

  // ---- Dates ---------------------------------------------------------------

  // New-month date = header date. Old-month date = last day of prior month.
  function computeDates(headerDate) {
    var newD = new Date(headerDate.y, headerDate.m - 1, headerDate.d);
    var oldD = new Date(headerDate.y, headerDate.m - 1, 0); // day 0 => prior month end
    return { newDate: newD, oldDate: oldD };
  }

  function mmddyyyy(d) {
    var mm = String(d.getMonth() + 1).padStart(2, "0");
    var dd = String(d.getDate()).padStart(2, "0");
    return mm + dd + String(d.getFullYear());
  }

  function journalNo(d) {
    return "Payroll" + mmddyyyy(d);
  }

  // ---- Journal building ----------------------------------------------------

  // Build one import sheet ("new" uses col F, "old" uses col G).
  // classMap (optional) maps numeric class IDs -> QBO class names.
  function buildSheet(rows, bookName, master, which, dateObj, classMap) {
    var accountMap = (master.books && master.books[bookName]) || {};
    var jNo = journalNo(dateObj);
    var amtKey = which === "new" ? "newAmt" : "oldAmt";

    var lines = [];
    var unmapped = {}; // memo -> true
    var sumDebit = 0;
    var sumCredit = 0; // non-clearing credits
    var clearingBefore = 0;

    rows.forEach(function (r) {
      var amt = round2(r[amtKey]);
      if (r.memo === CLEARING_MEMO) {
        clearingBefore = amt;
        return;
      }
      if (amt === 0) return; // zero lines have no journal effect

      var account = accountMap[r.memo];
      var isUnmapped = account === undefined;
      if (isUnmapped) {
        unmapped[r.memo] = true;
        account = "";
      }

      var debit = 0;
      var credit = 0;
      if (amt < 0) debit = round2(-amt);
      else credit = amt; // positive non-clearing amount -> credit

      sumDebit = round2(sumDebit + debit);
      sumCredit = round2(sumCredit + credit);

      lines.push({
        journalNo: jNo,
        journalDate: dateObj,
        memo: r.memo,
        account: account,
        debit: debit,
        credit: credit,
        description: r.memo,
        cls: resolveClass(r.cls, classMap),
        unmapped: isUnmapped,
      });
    });

    // Clearing wash line: set its credit so the journal balances exactly,
    // absorbing any penny rounding from the debit lines.
    var clearingCredit = round2(sumDebit - sumCredit);
    var clearingAccount = accountMap[CLEARING_MEMO];
    var clearingUnmapped = clearingAccount === undefined;
    if (clearingUnmapped) {
      unmapped[CLEARING_MEMO] = true;
      clearingAccount = "";
    }
    if (clearingCredit !== 0) {
      lines.push({
        journalNo: jNo,
        journalDate: dateObj,
        memo: CLEARING_MEMO,
        account: clearingAccount,
        debit: 0,
        credit: clearingCredit,
        description: CLEARING_MEMO,
        cls: "",
        unmapped: clearingUnmapped,
        isClearing: true,
      });
    }

    var totalDebit = round2(sumDebit);
    var totalCredit = round2(sumCredit + clearingCredit);

    return {
      which: which,
      journalNo: jNo,
      journalDate: dateObj,
      lines: lines,
      summary: {
        lineCount: lines.length,
        totalDebit: totalDebit,
        totalCredit: totalCredit,
        difference: round2(totalDebit - totalCredit),
        clearingBefore: clearingBefore,
        clearingAfter: clearingCredit,
        adjustment: round2(clearingCredit - clearingBefore),
        unmapped: Object.keys(unmapped),
      },
    };
  }

  function resolveClass(cls, classMap) {
    if (isBlank(cls)) return "";
    if (classMap && classMap[String(cls).trim()] !== undefined)
      return classMap[String(cls).trim()];
    return cls; // keep the raw ID
  }

  // Detect distinct numeric class IDs present in the input (for the mapping UI).
  function collectClassIds(rows) {
    var seen = {};
    var out = [];
    rows.forEach(function (r) {
      if (isBlank(r.cls)) return;
      var key = String(r.cls).trim();
      if (/^\d+$/.test(key) && !seen[key]) {
        seen[key] = true;
        out.push(key);
      }
    });
    return out;
  }

  // ---- Workbook generation -------------------------------------------------

  var COLS = [
    "Journal No",
    "Journal Date",
    "Memo",
    "Account",
    "Debit",
    "Credit",
    "Description",
    "Class",
  ];

  function headerStyle() {
    return {
      font: { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      fill: { patternType: "solid", fgColor: { rgb: HEADER_FILL } },
      alignment: { horizontal: "center", vertical: "center" },
      border: thinBorder("BFBFBF"),
    };
  }

  function thinBorder(rgb) {
    var s = { style: "thin", color: { rgb: rgb } };
    return { top: s, bottom: s, left: s, right: s };
  }

  function cell(v, style, type) {
    var c = { v: v, s: style };
    if (type) c.t = type;
    return c;
  }

  // Build a styled worksheet for one journal sheet.
  function makeJournalSheet(XLSX, sheet) {
    var rows = [];
    // header
    rows.push(
      COLS.map(function (h) {
        return cell(h, headerStyle(), "s");
      })
    );

    sheet.lines.forEach(function (ln, idx) {
      var alt = idx % 2 === 1;
      var base = {
        font: { name: "Arial", sz: 10 },
        fill: alt
          ? { patternType: "solid", fgColor: { rgb: ALT_FILL } }
          : undefined,
        border: thinBorder("D9D9D9"),
      };
      var money = Object.assign({}, base, {
        numFmt: CURRENCY_FMT,
        alignment: { horizontal: "right" },
      });
      var dateStyle = Object.assign({}, base, {
        numFmt: DATE_FMT,
        alignment: { horizontal: "center" },
      });
      var accStyle = ln.unmapped
        ? Object.assign({}, base, {
            fill: { patternType: "solid", fgColor: { rgb: UNMAPPED_FILL } },
          })
        : base;

      rows.push([
        cell(ln.journalNo, base, "s"),
        cell(ln.journalDate, dateStyle, "d"),
        cell(ln.memo, base, "s"),
        cell(ln.account, accStyle, "s"),
        ln.debit ? cell(ln.debit, money, "n") : cell("", money, "s"),
        ln.credit ? cell(ln.credit, money, "n") : cell("", money, "s"),
        cell(ln.description, base, "s"),
        isBlank(ln.cls) ? cell("", base, "s") : cell(ln.cls, base),
      ]);
    });

    var ws = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
    ws["!cols"] = [
      { wch: 16 },
      { wch: 13 },
      { wch: 30 },
      { wch: 52 },
      { wch: 14 },
      { wch: 14 },
      { wch: 30 },
      { wch: 10 },
    ];
    ws["!autofilter"] = { ref: "A1:H1" };
    return ws;
  }

  // Build the Summary sheet (first tab). Totals are formula-driven references
  // into the month sheets so they always match what QuickBooks will import.
  function makeSummarySheet(XLSX, newSheet, oldSheet, meta) {
    var navyTitle = {
      font: { name: "Arial", sz: 14, bold: true, color: { rgb: "FFFFFF" } },
      fill: { patternType: "solid", fgColor: { rgb: HEADER_FILL } },
      alignment: { horizontal: "left", vertical: "center" },
    };
    var label = { font: { name: "Arial", sz: 10, bold: true } };
    var val = { font: { name: "Arial", sz: 10 } };
    var money = {
      font: { name: "Arial", sz: 10 },
      numFmt: CURRENCY_FMT,
      alignment: { horizontal: "right" },
    };
    var hdr = headerStyle();

    var rows = [];
    rows.push([cell("Payroll → QuickBooks Journal Import — Summary", navyTitle, "s")]);
    rows.push([cell("", val, "s")]);
    rows.push([
      cell("Book / GL:", label, "s"),
      cell(meta.book, val, "s"),
    ]);
    rows.push([
      cell("Generated:", label, "s"),
      cell(meta.generated || "", val, "s"),
    ]);
    rows.push([cell("", val, "s")]);

    // Table header
    var tableHeaderRow = rows.length; // 0-based index of the header row
    rows.push(
      ["Journal No", "Journal Date", "Line Count", "Total Debit", "Total Credit", "Difference", "Status"].map(
        function (h) {
          return cell(h, hdr, "s");
        }
      )
    );

    // Row ranges in the month sheets (data starts at row 2; +1 for header).
    var newLast = newSheet.lines.length + 1;
    var oldLast = oldSheet.lines.length + 1;

    function monthRow(sheetName, s) {
      var last = s.lines.length + 1;
      var debitRef = "SUM('" + sheetName + "'!E2:E" + last + ")";
      var creditRef = "SUM('" + sheetName + "'!F2:F" + last + ")";
      return [
        cell(s.journalNo, val, "s"),
        cell(s.journalDate, Object.assign({}, val, { numFmt: DATE_FMT, alignment: { horizontal: "center" } }), "d"),
        cell(s.lines.length, Object.assign({}, val, { alignment: { horizontal: "center" } }), "n"),
        { t: "n", f: debitRef, s: money },
        { t: "n", f: creditRef, s: money },
        { t: "n", f: null, s: money }, // difference, filled below
        { t: "s", f: null, s: val }, // status, filled below
      ];
    }

    var newRowIdx = rows.length; // 0-based
    var newRow = monthRow("New Month", newSheet);
    rows.push(newRow);
    var oldRowIdx = rows.length;
    var oldRow = monthRow("Old Month", oldSheet);
    rows.push(oldRow);

    // Fill difference + status formulas (need the 1-based summary row numbers).
    function finishRow(rowObjs, rIdx1) {
      // Difference = Debit - Credit  (cols D,E on this summary row)
      rowObjs[5].f = "D" + rIdx1 + "-E" + rIdx1;
      // Status text
      rowObjs[6].f =
        'IF(ROUND(F' + rIdx1 + ',2)=0,"BALANCED","OUT OF BALANCE")';
      rowObjs[6].s = {
        font: { name: "Arial", sz: 10, bold: true },
        alignment: { horizontal: "center" },
      };
    }
    finishRow(newRow, newRowIdx + 1);
    finishRow(oldRow, oldRowIdx + 1);

    rows.push([cell("", val, "s")]);

    // Rounding note
    rows.push([cell("Clearing rounding adjustment", label, "s")]);
    rows.push([
      cell("New Month", val, "s"),
      cell("before:", val, "s"),
      cell(newSheet.summary.clearingBefore, money, "n"),
      cell("after:", val, "s"),
      cell(newSheet.summary.clearingAfter, money, "n"),
      cell("adj:", val, "s"),
      cell(newSheet.summary.adjustment, money, "n"),
    ]);
    rows.push([
      cell("Old Month", val, "s"),
      cell("before:", val, "s"),
      cell(oldSheet.summary.clearingBefore, money, "n"),
      cell("after:", val, "s"),
      cell(oldSheet.summary.clearingAfter, money, "n"),
      cell("adj:", val, "s"),
      cell(oldSheet.summary.adjustment, money, "n"),
    ]);

    rows.push([cell("", val, "s")]);

    // Unmapped memos
    var allUnmapped = {};
    newSheet.summary.unmapped.concat(oldSheet.summary.unmapped).forEach(function (m) {
      allUnmapped[m] = true;
    });
    var unmappedList = Object.keys(allUnmapped);
    if (unmappedList.length) {
      rows.push([
        cell("Unmapped memos (no account in " + meta.book + " — resolve before import):", {
          font: { name: "Arial", sz: 10, bold: true, color: { rgb: "C00000" } },
        }, "s"),
      ]);
      unmappedList.forEach(function (m) {
        rows.push([
          cell(m, {
            font: { name: "Arial", sz: 10 },
            fill: { patternType: "solid", fgColor: { rgb: UNMAPPED_FILL } },
          }, "s"),
        ]);
      });
    } else {
      rows.push([
        cell("All memos mapped to accounts. No unmapped items.", {
          font: { name: "Arial", sz: 10, italic: true, color: { rgb: "375623" } },
        }, "s"),
      ]);
    }

    var ws = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
    ws["!cols"] = [
      { wch: 20 },
      { wch: 14 },
      { wch: 12 },
      { wch: 16 },
      { wch: 16 },
      { wch: 14 },
      { wch: 16 },
    ];
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }];
    return ws;
  }

  // Assemble the full workbook. Returns the XLSX workbook object.
  function generateWorkbook(XLSX, newSheet, oldSheet, meta) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      makeSummarySheet(XLSX, newSheet, oldSheet, meta),
      "Summary"
    );
    XLSX.utils.book_append_sheet(wb, makeJournalSheet(XLSX, newSheet), "New Month");
    XLSX.utils.book_append_sheet(wb, makeJournalSheet(XLSX, oldSheet), "Old Month");
    return wb;
  }

  // End-to-end convenience: from parsed inputs to a workbook + report.
  function convert(XLSX, opts) {
    // opts: { headerDate, rows, master, book, classMap, generated }
    var dates = computeDates(opts.headerDate);
    var newSheet = buildSheet(opts.rows, opts.book, opts.master, "new", dates.newDate, opts.classMap);
    var oldSheet = buildSheet(opts.rows, opts.book, opts.master, "old", dates.oldDate, opts.classMap);
    var meta = { book: opts.book, generated: opts.generated || "" };
    var wb = generateWorkbook(XLSX, newSheet, oldSheet, meta);
    return { workbook: wb, newSheet: newSheet, oldSheet: oldSheet, dates: dates };
  }

  return {
    round2: round2,
    parseInput: parseInput,
    parseMaster: parseMaster,
    detectBook: detectBook,
    computeDates: computeDates,
    journalNo: journalNo,
    mmddyyyy: mmddyyyy,
    buildSheet: buildSheet,
    collectClassIds: collectClassIds,
    generateWorkbook: generateWorkbook,
    convert: convert,
    CLEARING_MEMO: CLEARING_MEMO,
  };
});
