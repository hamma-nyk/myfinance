const { google } = require("googleapis");
const ExcelJS = require("exceljs");

const BOT_TOKEN = process.env.BOT_TOKEN;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const GOOGLE_SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
const GOOGLE_PRIVATE_KEY = (process.env.GOOGLE_PRIVATE_KEY || "").replace(
  /\\n/g,
  "\n",
);
const ALLOWED_CHAT_ID = process.env.ALLOWED_CHAT_ID;
const REPORT_SHEET = "Laporan";

const MONTH_NAMES = [
  "Januari",
  "Februari",
  "Maret",
  "April",
  "Mei",
  "Juni",
  "Juli",
  "Agustus",
  "September",
  "Oktober",
  "November",
  "Desember",
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

const DATA_HEADERS = [
  "Tanggal",
  "Jam",
  "Bulan",
  "Tahun",
  "Tipe",
  "Deskripsi",
  "Jumlah",
];

function rgb(c) {
  return { red: c.red, green: c.green, blue: c.blue };
}

function cellVal(v) {
  if (typeof v === "number") return { userEnteredValue: { numberValue: v } };
  return { userEnteredValue: { stringValue: String(v || "") } };
}

function styledCell(v, bg, fg, bold) {
  const cell = cellVal(v);
  cell.userEnteredFormat = {};
  if (bg) cell.userEnteredFormat.backgroundColor = rgb(bg);
  if (fg)
    cell.userEnteredFormat.textFormat = {
      foregroundColor: rgb(fg),
      bold: !!bold,
    };
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
    range: `'${name}'!A1:G1`,
  });
  if (!res.data.values || res.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${name}'!A1:G1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [DATA_HEADERS] },
    });
    const headerCells = DATA_HEADERS.map((h) =>
      styledCell(h, COLORS.headerBg, COLORS.headerFg, true),
    );
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            updateCells: {
              rows: [{ values: headerCells }],
              range: {
                sheetId,
                startRowIndex: 0,
                endRowIndex: 1,
                startColumnIndex: 0,
                endColumnIndex: 7,
              },
              fields: "userEnteredValue,userEnteredFormat",
            },
          },
          {
            setBasicFilter: {
              filter: {
                range: {
                  sheetId,
                  startRowIndex: 0,
                  startColumnIndex: 0,
                  endColumnIndex: 7,
                },
              },
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
      requests: [
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: rowIndex,
              endRowIndex: rowIndex + 1,
              startColumnIndex: 0,
              endColumnIndex: 7,
            },
            cell: { userEnteredFormat: { backgroundColor: rgb(bg) } },
            fields: "userEnteredFormat.backgroundColor",
          },
        },
      ],
    },
  });
}

async function sendMessage(chatId, text, opts = {}) {
  await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      chat_id: chatId,
      text,
      parse_mode: "HTML",
      ...opts,
    }),
  });
}

async function sendDocument(chatId, buffer, filename, caption) {
  const form = new FormData();
  form.append("chat_id", String(chatId));
  form.append("document", new Blob([buffer]), filename);
  if (caption) form.append("caption", caption);
  form.append("parse_mode", "HTML");
  await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendDocument`, {
    method: "POST",
    body: form,
  });
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

function formatRupiah(num) {
  return "Rp " + Math.abs(num).toLocaleString("id-ID");
}

function parseTransaction(text) {
  const match = text.match(/^([+-]?)\s*(\d[\d.,]*)\s+(.+)$/);
  if (!match) return null;
  const sign = match[1];
  const amount = parseInt(match[2].replace(/[.,]/g, ""), 10);
  const desc = match[3].trim();
  const type = sign === "-" ? "pengeluaran" : "pemasukan";
  return { type, amount: type === "pengeluaran" ? -amount : amount, desc };
}

function jakartaNow() {
  return new Date(
    new Date().toLocaleString("en-US", { timeZone: "Asia/Jakarta" }),
  );
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
  ];
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${name}'!A:G`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [values] },
  });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${name}'!A:G`,
  });
  const rowCount = (res.data.values || []).length;
  await styleDataRow(sheets, sheetId, rowCount - 1, tx.type);
  return { now, year };
}

async function getRowsFromYear(sheets, year) {
  const name = String(year);
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${name}'!A:G`,
    });
    return (res.data.values || []).slice(1);
  } catch (_) {
    return [];
  }
}

async function getRowsMultiYear(sheets) {
  const all = await getMeta(sheets);
  const yearSheets = all
    .filter((s) => /^\d{4}$/.test(s.properties.title))
    .map((s) => s.properties.title);
  let rows = [];
  for (const name of yearSheets.sort()) {
    try {
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${name}'!A:G`,
      });
      rows = rows.concat((res.data.values || []).slice(1));
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
      styledCell("", COLORS.totalBg),
      styledCell("", COLORS.totalBg),
      styledCell("TOTAL PEMASUKAN", COLORS.totalBg, null, true),
      styledCell(totalIn, COLORS.totalBg, null, true),
      styledCell("", COLORS.white),
      styledCell("", COLORS.totalBg),
      styledCell("", COLORS.totalBg),
      styledCell("TOTAL PENGELUARAN", COLORS.totalBg, null, true),
      styledCell(totalOut, COLORS.totalBg, null, true),
    ],
  });

  allRows.push({ values: Array(9).fill(styledCell("", COLORS.white)) });
  const selisihBg = selisih >= 0 ? COLORS.selisihPlusBg : COLORS.selisihMinusBg;
  allRows.push({
    values: [
      styledCell("", selisihBg),
      styledCell("", selisihBg),
      styledCell("SELISIH", selisihBg, null, true),
      styledCell(selisih, selisihBg, null, true),
      styledCell("", COLORS.white),
      styledCell("", selisihBg),
      styledCell("", selisihBg),
      styledCell("", selisihBg),
      styledCell("", selisihBg),
    ],
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          updateCells: {
            range: {
              sheetId: reportSheetId,
              startRowIndex: 0,
              startColumnIndex: 0,
            },
            fields: "userEnteredValue,userEnteredFormat",
          },
        },
        {
          updateCells: {
            rows: allRows,
            range: {
              sheetId: reportSheetId,
              startRowIndex: 0,
              startColumnIndex: 0,
            },
            fields: "userEnteredValue,userEnteredFormat",
          },
        },
        {
          updateDimensionProperties: {
            range: {
              sheetId: reportSheetId,
              dimension: "COLUMNS",
              startIndex: 0,
              endIndex: 9,
            },
            properties: { pixelSize: 140 },
            fields: "pixelSize",
          },
        },
      ],
    },
  });

  return {
    totalIn,
    totalOut,
    selisih,
    countIn: pemasukan.length,
    countOut: pengeluaran.length,
  };
}

function parseFilter(text) {
  const parts = text.replace(/^\/(laporan|bulan|ringkasan)\s*/, "").trim();
  if (!parts) return null;
  const ddmmyyyy = parts.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmmyyyy)
    return {
      type: "date",
      day: ddmmyyyy[1].padStart(2, "0"),
      month: ddmmyyyy[2].padStart(2, "0"),
      year: ddmmyyyy[3],
    };
  const mmyyyy = parts.match(/^(\d{1,2})\/(\d{4})$/);
  if (mmyyyy)
    return {
      type: "month",
      month: mmyyyy[1].padStart(2, "0"),
      year: mmyyyy[2],
    };
  const yyyy = parts.match(/^(\d{4})$/);
  if (yyyy) return { type: "year", year: yyyy[1] };
  return null;
}

function filterRows(rows, filter) {
  if (!filter) return rows;
  if (filter.type === "date") {
    const d = `${filter.day}/${filter.month}/${filter.year}`;
    return rows.filter((r) => r[0] === d);
  }
  if (filter.type === "month") {
    return rows.filter((r) => {
      const p = r[0].split("/");
      return p.length === 3 && p[1] === filter.month && p[2] === filter.year;
    });
  }
  if (filter.type === "year") {
    return rows.filter((r) => {
      const p = r[0].split("/");
      return p.length === 3 && p[2] === filter.year;
    });
  }
  return rows;
}

function filterLabel(filter) {
  if (!filter) {
    const now = jakartaNow();
    return `${MONTH_NAMES[now.getMonth()]} ${now.getFullYear()}`;
  }
  if (filter.type === "date")
    return `${filter.day}/${filter.month}/${filter.year}`;
  if (filter.type === "month")
    return `${MONTH_NAMES[parseInt(filter.month) - 1]} ${filter.year}`;
  if (filter.type === "year") return filter.year;
}

async function getFilteredRows(sheets, filter) {
  if (filter && filter.type === "year") {
    return filterRows(await getRowsFromYear(sheets, filter.year), filter);
  }
  if (filter && (filter.type === "date" || filter.type === "month")) {
    return filterRows(await getRowsFromYear(sheets, filter.year), filter);
  }
  const now = jakartaNow();
  const year = now.getFullYear();
  const rows = await getRowsFromYear(sheets, year);
  if (!filter) {
    const m = String(now.getMonth() + 1).padStart(2, "0");
    return rows.filter((r) => {
      const p = r[0].split("/");
      return p.length === 3 && p[1] === m && p[2] === String(year);
    });
  }
  return filterRows(rows, filter);
}

async function handleMessage(msg) {
  const chatId = msg.chat.id;
  const text = (msg.text || "").trim();

  if (!text) return;

  if (ALLOWED_CHAT_ID && String(chatId) !== ALLOWED_CHAT_ID) {
    return sendMessage(chatId, "Maaf, bot ini hanya untuk pemiliknya.");
  }

  if (text === "/start") {
    return sendMessage(
      chatId,
      `<b>📒 MyFinance - Bot Catatan Finansial</b>\n\n` +
        `<b>Cara pakai:</b>\n` +
        `<blockquote>` +
        `+50000 gaji → pemasukan\n` +
        `-20000 kopi → pengeluaran\n` +
        `50000 freelance → pemasukan` +
        `</blockquote>\n` +
        `<b>Perintah:</b>\n` +
        `<blockquote>` +
        `/ringkasan → hari ini\n` +
        `/bulan → bulan ini\n` +
        `/laporan [filter] → laporan detail\n` +
        `/export [filter] → export .xlsx\n` +
        `/riwayat → 10 terakhir\n` +
        `/hapus → hapus terakhir` +
        `</blockquote>\n` +
        `<b>Filter:</b>\n` +
        `<blockquote>` +
        `/laporan 21/10/2026 → tanggal\n` +
        `/laporan 10/2026 → bulan\n` +
        `/laporan 2026 → tahun` +
        `</blockquote>`,
    );
  }

  if (text === "/help") {
    return sendMessage(
      chatId,
      `<b>Format transaksi:</b>\n` +
        `<blockquote>` +
        `+50000 gaji → pemasukan\n` +
        `-20000 kopi → pengeluaran\n` +
        `50000 freelance → pemasukan` +
        `</blockquote>\n` +
        `<b>Perintah:</b>\n` +
        `<blockquote>` +
        `/ringkasan → hari ini\n` +
        `/bulan → bulan ini\n` +
        `/laporan [filter] → laporan detail\n` +
        `/export [filter] → export .xlsx\n` +
        `/riwayat → 10 terakhir\n` +
        `/hapus → hapus terakhir` +
        `</blockquote>\n` +
        `<b>Filter:</b>\n` +
        `<blockquote>` +
        `/laporan 21/10/2026 → tanggal\n` +
        `/laporan 10/2026 → bulan\n` +
        `/laporan 2026 → tahun` +
        `</blockquote>`,
    );
  }

  const sheets = await getSheets();

  if (text === "/ringkasan") {
    const now = jakartaNow();
    const today = dateStr(now);
    const rows = await getRowsFromYear(sheets, now.getFullYear());
    const filtered = rows.filter((r) => r[0] === today);
    const r = await buildReport(sheets, filtered, `Ringkasan ${today}`);
    return sendMessage(
      chatId,
      `<b>📊 Ringkasan Hari Ini (${today})</b>\n\n` +
        `💰 Pemasukan: ${formatRupiah(r.totalIn)} (${r.countIn} transaksi)\n` +
        `💸 Pengeluaran: ${formatRupiah(r.totalOut)} (${r.countOut} transaksi)\n` +
        `📌 Selisih: ${r.selisih >= 0 ? "+" : "-"}${formatRupiah(r.selisih)}\n\n` +
        `📄 Lihat detail di sheet <b>${REPORT_SHEET}</b>`,
    );
  }

  if (text === "/bulan" || text.startsWith("/bulan ")) {
    const filter = parseFilter(text);
    const label = filterLabel(filter);
    const filtered = await getFilteredRows(sheets, filter);
    const r = await buildReport(sheets, filtered, `Ringkasan ${label}`);
    return sendMessage(
      chatId,
      `<b>📊 Ringkasan ${label}</b>\n\n` +
        `💰 Pemasukan: ${formatRupiah(r.totalIn)} (${r.countIn} transaksi)\n` +
        `💸 Pengeluaran: ${formatRupiah(r.totalOut)} (${r.countOut} transaksi)\n` +
        `📌 Selisih: ${r.selisih >= 0 ? "+" : "-"}${formatRupiah(r.selisih)}\n\n` +
        `📄 Lihat detail di sheet <b>${REPORT_SHEET}</b>`,
    );
  }

  if (text === "/laporan" || text.startsWith("/laporan ")) {
    const filter = parseFilter(text);
    const label = filterLabel(filter);
    const filtered = await getFilteredRows(sheets, filter);
    const r = await buildReport(sheets, filtered, `Laporan ${label}`);
    return sendMessage(
      chatId,
      `<b>📊 Laporan ${label}</b>\n\n` +
        `💰 Pemasukan: ${formatRupiah(r.totalIn)} (${r.countIn} transaksi)\n` +
        `💸 Pengeluaran: ${formatRupiah(r.totalOut)} (${r.countOut} transaksi)\n` +
        `📌 Selisih: ${r.selisih >= 0 ? "+" : "-"}${formatRupiah(r.selisih)}\n\n` +
        `📄 Lihat detail di sheet <b>${REPORT_SHEET}</b>`,
    );
  }

  if (text === "/export" || text.startsWith("/export ")) {
    const filter = parseFilter(text.replace("/export", "/laporan"));
    const label = filterLabel(filter);
    const filtered = await getFilteredRows(sheets, filter);
    if (filtered.length === 0) {
      return sendMessage(chatId, "Tidak ada data untuk diexport.");
    }
    const buffer = await exportExcel(filtered, `Laporan ${label}`);
    const filename = `laporan_${label.replace(/\//g, "-").replace(/\s/g, "_")}.xlsx`;
    return sendDocument(chatId, buffer, filename, `📊 <b>Export ${label}</b>\n${filtered.length} transaksi`);
  }

  if (text === "/riwayat") {
    const now = jakartaNow();
    const rows = await getRowsFromYear(sheets, now.getFullYear());
    const last10 = rows.slice(-10);
    if (last10.length === 0) {
      return sendMessage(chatId, "Belum ada transaksi.");
    }
    const lines = last10.map(
      (r) =>
        `${r[0]} ${r[1]} | ${r[4] === "pemasukan" ? "+" : "-"}${formatRupiah(Math.abs(Number(r[6])))} | ${r[5]}`,
    );
    return sendMessage(
      chatId,
      `<b>📋 10 Transaksi Terakhir</b>\n\n<pre>${lines.join("\n")}</pre>`,
    );
  }

  if (text === "/hapus") {
    const now = jakartaNow();
    const year = String(now.getFullYear());
    let yearSheetId;
    try {
      const all = await getMeta(sheets);
      const found = all.find((s) => s.properties.title === year);
      if (!found)
        return sendMessage(chatId, "Tidak ada transaksi untuk dihapus.");
      yearSheetId = found.properties.sheetId;
    } catch (_) {
      return sendMessage(chatId, "Tidak ada transaksi untuk dihapus.");
    }
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${year}'!A:G`,
    });
    const allRows = res.data.values || [];
    if (allRows.length < 2) {
      return sendMessage(chatId, "Tidak ada transaksi untuk dihapus.");
    }
    const lastIdx = allRows.length - 1;
    const deleted = allRows[lastIdx];
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            deleteDimension: {
              range: {
                sheetId: yearSheetId,
                dimension: "ROWS",
                startIndex: lastIdx,
                endIndex: lastIdx + 1,
              },
            },
          },
        ],
      },
    });
    return sendMessage(
      chatId,
      `🗑 Dihapus: ${deleted[5]} (${deleted[4] === "pemasukan" ? "+" : "-"}${formatRupiah(Math.abs(Number(deleted[6])))})`,
    );
  }

  const tx = parseTransaction(text);
  if (!tx) {
    return sendMessage(
      chatId,
      "Format tidak dikenali. Contoh:\n<code>+50000 gaji</code>\n<code>-20000 kopi</code>\n\nKetik /help untuk bantuan.",
    );
  }

  const { now } = await appendTransaction(sheets, chatId, tx);
  const today = dateStr(now);
  const rows = await getRowsFromYear(sheets, now.getFullYear());
  const todayRows = rows.filter((r) => r[0] === today);
  const r = await buildReport(sheets, todayRows, `Laporan Hari Ini (${today})`);

  const emoji = tx.type === "pemasukan" ? "💰" : "💸";
  return sendMessage(
    chatId,
    `${emoji} <b>${tx.type.charAt(0).toUpperCase() + tx.type.slice(1)} tercatat!</b>\n\n` +
      `📝 ${tx.desc}\n` +
      `💵 ${tx.amount > 0 ? "+" : "-"}${formatRupiah(tx.amount)}\n` +
      `🕐 ${dateStr(now)} ${timeStr(now)}\n\n` +
      `📊 Hari ini: +${formatRupiah(r.totalIn)} / -${formatRupiah(r.totalOut)} = ${r.selisih >= 0 ? "+" : "-"}${formatRupiah(r.selisih)}`,
  );
}

async function poll(offset = 0) {
  try {
    const res = await fetch(
      `https://api.telegram.org/bot${BOT_TOKEN}/getUpdates?offset=${offset}&timeout=30`,
    );
    const data = await res.json();
    if (data.ok && data.result.length > 0) {
      for (const update of data.result) {
        if (update.message) {
          console.log(
            `[${update.message.chat.id}] ${update.message.text || ""}`,
          );
          await handleMessage(update.message).catch((e) =>
            console.error("handleMessage error:", e.message),
          );
        }
        offset = update.update_id + 1;
      }
    }
  } catch (e) {
    console.error("Poll error:", e.message);
    await new Promise((r) => setTimeout(r, 3000));
  }
  poll(offset);
}

console.log("MyFinance bot running in polling mode...");
fetch(`https://api.telegram.org/bot${BOT_TOKEN}/deleteWebhook`)
  .then(() => poll())
  .catch((e) => {
    console.error("Failed to start:", e.message);
    process.exit(1);
  });
