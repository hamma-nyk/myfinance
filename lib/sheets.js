const { google } = require("googleapis");
const ExcelJS = require("exceljs");

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const GOOGLE_SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
const GOOGLE_PRIVATE_KEY = (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n");
const REPORT_SHEET = "Laporan";

const MONTH_NAMES = [
  "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  "Juli", "Agustus", "September", "Oktober", "November", "Desember",
];

const COLORS = {
  headerBg: { red: 0.16, green: 0.25, blue: 0.47 },
  headerFg: { red: 1, green: 1, blue: 1 },
  incomeBg: { red: 0.85, green: 0.95, blue: 0.85 },
  expenseBg: { red: 0.97, green: 0.85, blue: 0.85 },
  totalBg: { red: 0.93, green: 0.93, blue: 0.93 },
  selisihPlusBg: { red: 0.72, green: 0.88, blue: 0.72 },
  selisihMinusBg: { red: 0.92, green: 0.72, blue: 0.72 },
  white: { red: 1, green: 1, blue: 1 },
};

const DATA_HEADERS = ["Tanggal", "Jam", "Bulan", "Tahun", "Tipe", "Deskripsi", "Jumlah", "Chat ID"];

function rgb(c) { return { red: c.red, green: c.green, blue: c.blue }; }

function cellVal(v) {
  if (typeof v === "number") return { userEnteredValue: { numberValue: v } };
  return { userEnteredValue: { stringValue: String(v || "") } };
}

function styledCell(v, bg, fg, bold) {
  const cell = cellVal(v);
  cell.userEnteredFormat = {};
  if (bg) cell.userEnteredFormat.backgroundColor = rgb(bg);
  if (fg) cell.userEnteredFormat.textFormat = { foregroundColor: rgb(fg), bold: !!bold };
  else if (bold) cell.userEnteredFormat.textFormat = { bold: true };
  return cell;
}

async function getSheets() {
  const auth = new google.auth.JWT({
    email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: GOOGLE_PRIVATE_KEY,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

async function getMeta(sheets) {
  const res = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  return res.data.sheets;
}

async function ensureSheet(sheets, name) {
  const all = await getMeta(sheets);
  const found = all.find((s) => s.properties.title === name);
  if (found) return found.properties.sheetId;
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [{ addSheet: { properties: { title: name } } }] },
  });
  return res.data.replies[0].addSheet.properties.sheetId;
}

async function ensureYearSheet(sheets, year) {
  const name = String(year);
  const sheetId = await ensureSheet(sheets, name);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${name}'!A1:H1`,
  });
  if (!res.data.values || res.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${name}'!A1:H1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [DATA_HEADERS] },
    });
    const headerCells = DATA_HEADERS.map((h) => styledCell(h, COLORS.headerBg, COLORS.headerFg, true));
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            updateCells: {
              rows: [{ values: headerCells }],
              range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 8 },
              fields: "userEnteredValue,userEnteredFormat",
            },
          },
          {
            setBasicFilter: {
              filter: { range: { sheetId, startRowIndex: 0, startColumnIndex: 0, endColumnIndex: 8 } },
            },
          },
        ],
      },
    });
  }
  return { sheetId, name };
}

async function styleDataRow(sheets, sheetId, rowIndex, type) {
  const bg = type === "pemasukan" ? COLORS.incomeBg : COLORS.expenseBg;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{
        repeatCell: {
          range: { sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1, startColumnIndex: 0, endColumnIndex: 8 },
          cell: { userEnteredFormat: { backgroundColor: rgb(bg) } },
          fields: "userEnteredFormat.backgroundColor",
        },
      }],
    },
  });
}

function jakartaNow() {
  return new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Jakarta" }));
}

function dateStr(d) {
  return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
}

function timeStr(d) {
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}:${String(d.getSeconds()).padStart(2, "0")}`;
}

async function appendTransaction(sheets, chatId, tx) {
  const now = jakartaNow();
  const year = now.getFullYear();
  const { sheetId, name } = await ensureYearSheet(sheets, year);
  const values = [
    dateStr(now),
    timeStr(now),
    MONTH_NAMES[now.getMonth()],
    year,
    tx.type,
    tx.desc,
    tx.amount,
    String(chatId),
  ];
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${name}'!A:H`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [values] },
  });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${name}'!A:H`,
  });
  const rowCount = (res.data.values || []).length;
  await styleDataRow(sheets, sheetId, rowCount - 1, tx.type);
  return { now, year };
}

function matchChatId(row, chatId) {
  if (!chatId) return true;
  return !row[7] || String(row[7]) === String(chatId);
}

async function getRowsFromYear(sheets, year, chatId) {
  const name = String(year);
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${name}'!A:H`,
    });
    const rows = (res.data.values || []).slice(1);
    return rows.filter((r) => matchChatId(r, chatId));
  } catch (_) {
    return [];
  }
}

async function getRowsMultiYear(sheets, chatId) {
  const all = await getMeta(sheets);
  const yearSheets = all.filter((s) => /^\d{4}$/.test(s.properties.title)).map((s) => s.properties.title);
  let rows = [];
  for (const name of yearSheets.sort()) {
    try {
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${name}'!A:H`,
      });
      const yrRows = (res.data.values || []).slice(1);
      rows = rows.concat(yrRows.filter((r) => matchChatId(r, chatId)));
    } catch (_) {}
  }
  return rows;
}

async function buildReport(sheets, rows, title) {
  const pemasukan = rows.filter((r) => r[4] === "pemasukan");
  const pengeluaran = rows.filter((r) => r[4] === "pengeluaran");
  const maxRows = Math.max(pemasukan.length, pengeluaran.length, 1);

  const reportSheetId = await ensureSheet(sheets, REPORT_SHEET);

  const allRows = [];
  allRows.push({
    values: [
      styledCell(title, COLORS.headerBg, COLORS.headerFg, true),
      ...Array(8).fill(styledCell("", COLORS.headerBg)),
    ],
  });
  allRows.push({ values: Array(9).fill(styledCell("", COLORS.white)) });
  allRows.push({
    values: [
      styledCell("Tanggal", COLORS.headerBg, COLORS.headerFg, true),
      styledCell("Jam", COLORS.headerBg, COLORS.headerFg, true),
      styledCell("Pemasukan", COLORS.headerBg, COLORS.headerFg, true),
      styledCell("Jumlah", COLORS.headerBg, COLORS.headerFg, true),
      styledCell("", COLORS.white),
      styledCell("Tanggal", COLORS.headerBg, COLORS.headerFg, true),
      styledCell("Jam", COLORS.headerBg, COLORS.headerFg, true),
      styledCell("Pengeluaran", COLORS.headerBg, COLORS.headerFg, true),
      styledCell("Jumlah", COLORS.headerBg, COLORS.headerFg, true),
    ],
  });

  for (let i = 0; i < maxRows; i++) {
    const p = pemasukan[i];
    const e = pengeluaran[i];
    allRows.push({
      values: [
        styledCell(p ? p[0] : "", COLORS.incomeBg),
        styledCell(p ? p[1] : "", COLORS.incomeBg),
        styledCell(p ? p[5] : "", COLORS.incomeBg),
        styledCell(p ? Number(p[6]) : "", COLORS.incomeBg),
        styledCell("", COLORS.white),
        styledCell(e ? e[0] : "", COLORS.expenseBg),
        styledCell(e ? e[1] : "", COLORS.expenseBg),
        styledCell(e ? e[5] : "", COLORS.expenseBg),
        styledCell(e ? Math.abs(Number(e[6])) : "", COLORS.expenseBg),
      ],
    });
  }

  const totalIn = pemasukan.reduce((s, r) => s + Number(r[6]), 0);
  const totalOut = pengeluaran.reduce((s, r) => s + Math.abs(Number(r[6])), 0);
  const selisih = totalIn - totalOut;

  allRows.push({ values: Array(9).fill(styledCell("", COLORS.white)) });
  allRows.push({
    values: [
      styledCell("", COLORS.totalBg), styledCell("", COLORS.totalBg),
      styledCell("TOTAL PEMASUKAN", COLORS.totalBg, null, true),
      styledCell(totalIn, COLORS.totalBg, null, true),
      styledCell("", COLORS.white),
      styledCell("", COLORS.totalBg), styledCell("", COLORS.totalBg),
      styledCell("TOTAL PENGELUARAN", COLORS.totalBg, null, true),
      styledCell(totalOut, COLORS.totalBg, null, true),
    ],
  });

  allRows.push({ values: Array(9).fill(styledCell("", COLORS.white)) });
  const selisihBg = selisih >= 0 ? COLORS.selisihPlusBg : COLORS.selisihMinusBg;
  allRows.push({
    values: [
      styledCell("", selisihBg), styledCell("", selisihBg),
      styledCell("SELISIH", selisihBg, null, true),
      styledCell(selisih, selisihBg, null, true),
      styledCell("", COLORS.white),
      styledCell("", selisihBg), styledCell("", selisihBg),
      styledCell("", selisihBg), styledCell("", selisihBg),
    ],
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          updateCells: {
            range: { sheetId: reportSheetId, startRowIndex: 0, startColumnIndex: 0 },
            fields: "userEnteredValue,userEnteredFormat",
          },
        },
        {
          updateCells: {
            rows: allRows,
            range: { sheetId: reportSheetId, startRowIndex: 0, startColumnIndex: 0 },
            fields: "userEnteredValue,userEnteredFormat",
          },
        },
        {
          updateDimensionProperties: {
            range: { sheetId: reportSheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 9 },
            properties: { pixelSize: 140 },
            fields: "pixelSize",
          },
        },
      ],
    },
  });

  return { totalIn, totalOut, selisih, countIn: pemasukan.length, countOut: pengeluaran.length };
}

async function exportExcel(rows, title) {
  const wb = new ExcelJS.Workbook();
  const wsData = wb.addWorksheet("Data");
  wsData.columns = [
    { header: "Tanggal", key: "tgl", width: 14 },
    { header: "Jam", key: "jam", width: 10 },
    { header: "Bulan", key: "bln", width: 12 },
    { header: "Tahun", key: "thn", width: 8 },
    { header: "Tipe", key: "tipe", width: 12 },
    { header: "Deskripsi", key: "desc", width: 25 },
    { header: "Jumlah", key: "jml", width: 15 },
  ];
  const hdrFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF293F78" } };
  const hdrFont = { color: { argb: "FFFFFFFF" }, bold: true };
  wsData.getRow(1).eachCell((c) => { c.fill = hdrFill; c.font = hdrFont; });
  rows.forEach((r) => {
    const row = wsData.addRow({ tgl: r[0], jam: r[1], bln: r[2], thn: r[3], tipe: r[4], desc: r[5], jml: Number(r[6]) });
    const bg = r[4] === "pemasukan" ? "FFD9F2D9" : "FFF8D9D9";
    row.eachCell((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } }; });
  });

  const pemasukan = rows.filter((r) => r[4] === "pemasukan");
  const pengeluaran = rows.filter((r) => r[4] === "pengeluaran");
  const totalIn = pemasukan.reduce((s, r) => s + Number(r[6]), 0);
  const totalOut = pengeluaran.reduce((s, r) => s + Math.abs(Number(r[6])), 0);
  const selisih = totalIn - totalOut;
  const maxRows = Math.max(pemasukan.length, pengeluaran.length, 1);

  const wsReport = wb.addWorksheet("Laporan");
  wsReport.columns = [
    { width: 14 }, { width: 10 }, { width: 20 }, { width: 15 },
    { width: 3 },
    { width: 14 }, { width: 10 }, { width: 20 }, { width: 15 },
  ];
  const titleRow = wsReport.addRow([title]);
  titleRow.eachCell((c) => { c.fill = hdrFill; c.font = hdrFont; });
  wsReport.mergeCells(1, 1, 1, 9);
  wsReport.addRow([]);
  const hdr2 = wsReport.addRow(["Tanggal", "Jam", "Pemasukan", "Jumlah", "", "Tanggal", "Jam", "Pengeluaran", "Jumlah"]);
  hdr2.eachCell((c, i) => { if (i !== 5) { c.fill = hdrFill; c.font = hdrFont; } });

  for (let i = 0; i < maxRows; i++) {
    const p = pemasukan[i];
    const e = pengeluaran[i];
    const row = wsReport.addRow([
      p ? p[0] : "", p ? p[1] : "", p ? p[5] : "", p ? Number(p[6]) : "",
      "",
      e ? e[0] : "", e ? e[1] : "", e ? e[5] : "", e ? Math.abs(Number(e[6])) : "",
    ]);
    row.eachCell((c, i) => {
      if (i <= 4 && i !== 5) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9F2D9" } };
      if (i >= 6) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8D9D9" } };
    });
  }

  wsReport.addRow([]);
  const totRow = wsReport.addRow(["", "", "TOTAL PEMASUKAN", totalIn, "", "", "", "TOTAL PENGELUARAN", totalOut]);
  const totFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEDEDED" } };
  totRow.eachCell((c, i) => { if (i !== 5) { c.fill = totFill; c.font = { bold: true }; } });

  wsReport.addRow([]);
  const selRow = wsReport.addRow(["", "", "SELISIH", selisih]);
  const selFill = { type: "pattern", pattern: "solid", fgColor: { argb: selisih >= 0 ? "FFB8E0B8" : "FFEBB8B8" } };
  selRow.eachCell((c) => { c.fill = selFill; c.font = { bold: true }; });

  return Buffer.from(await wb.xlsx.writeBuffer());
}

function computeStats(rows) {
  const pemasukan = rows.filter((r) => r[4] === "pemasukan");
  const pengeluaran = rows.filter((r) => r[4] === "pengeluaran");

  const totalIncome = pemasukan.reduce((s, r) => s + Number(r[6]), 0);
  const totalExpense = pengeluaran.reduce((s, r) => s + Math.abs(Number(r[6])), 0);
  const net = totalIncome - totalExpense;
  const savingsRate = totalIncome > 0 ? Math.round(((totalIncome - totalExpense) / totalIncome) * 100) : 0;

  // Category breakdown for expenses
  const categoryMap = {};
  pengeluaran.forEach((r) => {
    const rawCat = (r[5] || "Lainnya").trim().toLowerCase();
    const cat = rawCat.charAt(0).toUpperCase() + rawCat.slice(1);
    const amt = Math.abs(Number(r[6]));
    categoryMap[cat] = (categoryMap[cat] || 0) + amt;
  });

  const sortedCategories = Object.entries(categoryMap)
    .map(([name, amount]) => ({ name, amount }))
    .sort((a, b) => b.amount - a.amount);

  const topCategories = sortedCategories.slice(0, 7);
  if (sortedCategories.length > 7) {
    const otherAmount = sortedCategories.slice(7).reduce((s, c) => s + c.amount, 0);
    topCategories.push({ name: "Lainnya", amount: otherAmount });
  }

  // Daily or Timeline trend (group by date)
  const timelineMap = {};
  rows.forEach((r) => {
    const date = r[0]; // DD/MM/YYYY
    if (!timelineMap[date]) timelineMap[date] = { date, income: 0, expense: 0 };
    const amt = Number(r[6]);
    if (r[4] === "pemasukan") timelineMap[date].income += amt;
    else timelineMap[date].expense += Math.abs(amt);
  });

  const timeline = Object.values(timelineMap);

  // Recent transactions (last 20, newest first)
  const recent = rows
    .slice(-20)
    .reverse()
    .map((r) => ({
      date: r[0],
      time: r[1],
      type: r[4],
      desc: r[5],
      amount: Number(r[6]),
    }));

  return {
    summary: {
      totalIncome,
      totalExpense,
      net,
      savingsRate,
      txCount: rows.length,
      incomeCount: pemasukan.length,
      expenseCount: pengeluaran.length,
    },
    categories: topCategories,
    timeline,
    recent,
  };
}

module.exports = {
  MONTH_NAMES,
  REPORT_SHEET,
  getSheets,
  getMeta,
  ensureYearSheet,
  appendTransaction,
  getRowsFromYear,
  getRowsMultiYear,
  buildReport,
  exportExcel,
  computeStats,
  jakartaNow,
  dateStr,
  timeStr,
};
