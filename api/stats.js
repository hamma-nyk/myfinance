const { verifyAccessToken } = require("../lib/auth");
const { getSheets, getRowsMultiYear, computeStats, jakartaNow } = require("../lib/sheets");

function parseRowDate(str) {
  if (!str) return null;
  const parts = String(str).trim().split(/[\/\-]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      return { year: parseInt(parts[0], 10), month: parseInt(parts[1], 10), day: parseInt(parts[2], 10) };
    }
    return { day: parseInt(parts[0], 10), month: parseInt(parts[1], 10), year: parseInt(parts[2], 10) };
  }
  return null;
}

function filterByPeriod(rows, period, now) {
  const curYear = now.getFullYear();
  const curMonth = now.getMonth() + 1;
  const curDay = now.getDate();

  if (period === "today") {
    return rows.filter((r) => {
      const d = parseRowDate(r[0]);
      return d && d.day === curDay && d.month === curMonth && d.year === curYear;
    });
  }

  if (period === "month") {
    return rows.filter((r) => {
      const d = parseRowDate(r[0]);
      return d && d.month === curMonth && d.year === curYear;
    });
  }

  if (period === "3months") {
    const nowMonths = curYear * 12 + curMonth;
    return rows.filter((r) => {
      const d = parseRowDate(r[0]);
      if (!d) return false;
      const rowMonths = d.year * 12 + d.month;
      const diff = nowMonths - rowMonths;
      return diff >= 0 && diff < 3;
    });
  }

  if (period === "year") {
    return rows.filter((r) => {
      const d = parseRowDate(r[0]);
      return d && d.year === curYear;
    });
  }

  return rows;
}

module.exports = async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { token, period = "month" } = req.query || {};
  const auth = verifyAccessToken(token);

  if (!auth) {
    return res.status(401).json({
      error: "Token tidak valid atau sudah kadaluarsa. Kirim perintah /link di Telegram.",
    });
  }

  try {
    const sheets = await getSheets();
    const now = jakartaNow();
    const rows = await getRowsMultiYear(sheets, auth.chatId);
    const filtered = filterByPeriod(rows, period, now);
    const stats = computeStats(filtered);

    res.setHeader("Cache-Control", "no-store");
    return res.status(200).json({
      chatId: auth.chatId,
      period,
      ...stats,
    });
  } catch (err) {
    console.error("Stats API error:", err);
    return res.status(500).json({ error: "Gagal memproses data statistik" });
  }
};
