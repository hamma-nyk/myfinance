const assert = require("assert");
const { generateAccessToken, verifyAccessToken, isAllowed } = require("./lib/auth");
const { computeStats } = require("./lib/sheets");

process.env.ALLOWED_CHAT_ID = "111,222";
process.env.APP_SECRET = "test-secret-key";

// 1. Test whitelist
assert.strictEqual(isAllowed("111"), true);
assert.strictEqual(isAllowed("222"), true);
assert.strictEqual(isAllowed("333"), false);

// 2. Test token generation & verification
const token111 = generateAccessToken("111");
const verified111 = verifyAccessToken(token111);
assert(verified111 !== null, "Token should be valid");
assert.strictEqual(verified111.chatId, "111");

// 3. Test tamper resistance
const raw = Buffer.from(token111, "base64url").toString("utf8");
const tamperedRaw = raw.replace("111", "222"); // replace chatId without changing HMAC
const tamperedToken = Buffer.from(tamperedRaw).toString("base64url");
assert.strictEqual(verifyAccessToken(tamperedToken), null, "Tampered token must be rejected");

// 4. Test unauthorized user token rejected
const token333 = generateAccessToken("333");
assert.strictEqual(verifyAccessToken(token333), null, "Non-allowed user token must be rejected");

// 5. Test stats computation
const sampleRows = [
  ["21/10/2026", "10:00:00", "Oktober", 2026, "pemasukan", "Gaji", 5000000, "111"],
  ["21/10/2026", "11:00:00", "Oktober", 2026, "pengeluaran", "Kopi", -20000, "111"],
  ["21/10/2026", "12:00:00", "Oktober", 2026, "pengeluaran", "Makan", -30000, "111"],
];
const stats = computeStats(sampleRows);
assert.strictEqual(stats.summary.totalIncome, 5000000);
assert.strictEqual(stats.summary.totalExpense, 50000);
assert.strictEqual(stats.summary.net, 4950000);
assert.strictEqual(stats.categories.length, 2);
assert.strictEqual(stats.timeline.length, 1);

console.log("All unit tests passed successfully!");
