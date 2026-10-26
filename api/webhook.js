const { isAllowed, generateAccessToken } = require("../lib/auth");
const {
  MONTH_NAMES,
  REPORT_SHEET,
  getSheets,
  getMeta,
  appendTransaction,
  getRowsFromYear,
  getRowsMultiYear,
  buildReport,
  exportExcel,
  jakartaNow,
  dateStr,
  timeStr,
} = require("../lib/sheets");

const BOT_TOKEN = process.env.BOT_TOKEN;

async function sendMessage(chatId, text, opts = {}) {
  await fetch(`https://api.telegram.org/bot${BOT_TOKEN}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: chatId, text, parse_mode: "HTML", ...opts }),
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

function parseFilter(text) {
  const parts = text.replace(/^\/(laporan|bulan|ringkasan|export)\s*/, "").trim();
  if (!parts) return null;
  const ddmmyyyy = parts.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (ddmmyyyy) return { type: "date", day: ddmmyyyy[1].padStart(2, "0"), month: ddmmyyyy[2].padStart(2, "0"), year: ddmmyyyy[3] };
  const mmyyyy = parts.match(/^(\d{1,2})\/(\d{4})$/);
  if (mmyyyy) return { type: "month", month: mmyyyy[1].padStart(2, "0"), year: mmyyyy[2] };
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
      const p = (r[0] || "").split("/");
      return p.length === 3 && p[1] === filter.month && p[2] === filter.year;
    });
  }
  if (filter.type === "year") {
    return rows.filter((r) => {
      const p = (r[0] || "").split("/");
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
  if (filter.type === "date") return `${filter.day}/${filter.month}/${filter.year}`;
  if (filter.type === "month") return `${MONTH_NAMES[parseInt(filter.month) - 1]} ${filter.year}`;
  if (filter.type === "year") return filter.year;
}

async function getFilteredRows(sheets, filter, chatId) {
  if (filter && (filter.type === "date" || filter.type === "month" || filter.type === "year")) {
    return filterRows(await getRowsFromYear(sheets, filter.year, chatId), filter);
  }
  const now = jakartaNow();
  const year = now.getFullYear();
  const rows = await getRowsFromYear(sheets, year, chatId);
  if (!filter) {
    const m = String(now.getMonth() + 1).padStart(2, "0");
    return rows.filter((r) => {
      const p = (r[0] || "").split("/");
      return p.length === 3 && p[1] === m && p[2] === String(year);
    });
  }
  return filterRows(rows, filter);
}

async function handleMessage(msg, hostBaseUrl) {
  const chatId = msg.chat.id;
  const text = (msg.text || "").trim();

  if (!text) return;

  if (!isAllowed(chatId)) {
    return sendMessage(chatId, "Maaf, bot ini hanya untuk pengguna terdaftar.");
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
        `/link → buka web dashboard & statistik\n` +
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
        `/link → buka web dashboard & statistik\n` +
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

  if (text === "/link") {
    const token = generateAccessToken(chatId);
    const baseUrl = process.env.APP_URL || hostBaseUrl || "http://localhost:3000";
    const dashboardUrl = `${baseUrl.replace(/\/$/, "")}/?token=${token}`;
    return sendMessage(
      chatId,
      `<b>📊 Dashboard Finansial Pribadi</b>\n\n` +
        `Klik tautan di bawah untuk melihat diagram, statistik, dan riwayat:\n\n` +
        `<a href="${dashboardUrl}">${dashboardUrl}</a>\n\n` +
        `<blockquote>Tautan ini memuat token akses pribadi Anda dan berlaku 30 hari. Jangan bagikan ke orang lain.</blockquote>`,
    );
  }

  const sheets = await getSheets();

  if (text === "/ringkasan") {
    const now = jakartaNow();
    const today = dateStr(now);
    const rows = await getRowsFromYear(sheets, now.getFullYear(), chatId);
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
    const filtered = await getFilteredRows(sheets, filter, chatId);
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
    const filtered = await getFilteredRows(sheets, filter, chatId);
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
    const filter = parseFilter(text);
    const label = filterLabel(filter);
    const filtered = await getFilteredRows(sheets, filter, chatId);
    if (filtered.length === 0) {
      return sendMessage(chatId, "Tidak ada data untuk diexport.");
    }
    const buffer = await exportExcel(filtered, `Laporan ${label}`);
    const filename = `laporan_${label.replace(/\//g, "-").replace(/\s/g, "_")}.xlsx`;
    return sendDocument(chatId, buffer, filename, `📊 <b>Export ${label}</b>\n${filtered.length} transaksi`);
  }

  if (text === "/riwayat") {
    const now = jakartaNow();
    const rows = await getRowsFromYear(sheets, now.getFullYear(), chatId);
    const last10 = rows.slice(-10);
    if (last10.length === 0) {
      return sendMessage(chatId, "Belum ada transaksi.");
    }
    const lines = last10.map(
      (r) => `${r[0]} ${r[1]} | ${r[4] === "pemasukan" ? "+" : "-"}${formatRupiah(Math.abs(Number(r[6])))} | ${r[5]}`,
    );
    return sendMessage(chatId, `<b>📋 10 Transaksi Terakhir</b>\n\n<pre>${lines.join("\n")}</pre>`);
  }

  if (text === "/hapus") {
    const now = jakartaNow();
    const year = String(now.getFullYear());
    let yearSheetId;
    try {
      const all = await getMeta(sheets);
      const found = all.find((s) => s.properties.title === year);
      if (!found) return sendMessage(chatId, "Tidak ada transaksi untuk dihapus.");
      yearSheetId = found.properties.sheetId;
    } catch (_) {
      return sendMessage(chatId, "Tidak ada transaksi untuk dihapus.");
    }
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.SPREADSHEET_ID,
      range: `'${year}'!A:H`,
    });
    const allRows = res.data.values || [];
    let lastIdx = -1;
    for (let i = allRows.length - 1; i >= 1; i--) {
      if (!allRows[i][7] || String(allRows[i][7]) === String(chatId)) {
        lastIdx = i;
        break;
      }
    }
    if (lastIdx === -1) {
      return sendMessage(chatId, "Tidak ada transaksi untuk dihapus.");
    }
    const deleted = allRows[lastIdx];
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: process.env.SPREADSHEET_ID,
      requestBody: {
        requests: [{
          deleteDimension: {
            range: { sheetId: yearSheetId, dimension: "ROWS", startIndex: lastIdx, endIndex: lastIdx + 1 },
          },
        }],
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
  const rows = await getRowsFromYear(sheets, now.getFullYear(), chatId);
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

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(200).send("MyFinance Bot is running.");
  }
  try {
    const proto = req.headers["x-forwarded-proto"] || "https";
    const host = req.headers["x-forwarded-host"] || req.headers.host;
    const hostBaseUrl = host ? `${proto}://${host}` : undefined;

    const { message } = req.body || {};
    if (message) {
      await handleMessage(message, hostBaseUrl);
    }
  } catch (err) {
    console.error("Error:", err);
  }
  res.status(200).send("OK");
};

module.exports.handleMessage = handleMessage;
module.exports.sendMessage = sendMessage;
