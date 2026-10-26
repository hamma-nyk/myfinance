const crypto = require("crypto");

const SECRET = process.env.APP_SECRET || process.env.BOT_TOKEN || "catatai-secret-key";

function isAllowed(chatId) {
  const allowed = process.env.ALLOWED_CHAT_ID;
  if (!allowed) return true;
  const list = String(allowed).split(",").map((s) => s.trim());
  return list.includes(String(chatId));
}

function generateAccessToken(chatId, expiresInDays = 30) {
  const expiry = Date.now() + expiresInDays * 24 * 60 * 60 * 1000;
  const payload = `${chatId}.${expiry}`;
  const sig = crypto.createHmac("sha256", SECRET).update(payload).digest("hex");
  return Buffer.from(`${payload}.${sig}`).toString("base64url");
}

function verifyAccessToken(token) {
  if (!token) return null;
  try {
    const raw = Buffer.from(token, "base64url").toString("utf8");
    const [chatId, expiryStr, sig] = raw.split(".");
    if (!chatId || !expiryStr || !sig) return null;
    const expiry = Number(expiryStr);
    if (isNaN(expiry) || Date.now() > expiry) return null;
    const expectedSig = crypto.createHmac("sha256", SECRET).update(`${chatId}.${expiry}`).digest("hex");
    if (!crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expectedSig))) return null;
    if (!isAllowed(chatId)) return null;
    return { chatId, expiry };
  } catch (_) {
    return null;
  }
}

module.exports = { isAllowed, generateAccessToken, verifyAccessToken };
