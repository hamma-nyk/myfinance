const http = require("http");
const fs = require("fs");
const path = require("path");
const { handleMessage } = require("./api/webhook");
const statsHandler = require("./api/stats");

const BOT_TOKEN = process.env.BOT_TOKEN;
const PORT = process.env.PORT || 3000;
const HOST_BASE = process.env.APP_URL || `http://localhost:${PORT}`;

// 1. Local HTTP Server for Dashboard & Stats API
const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host}`);
  const pathname = parsedUrl.pathname;

  // API stats route
  if (pathname === "/api/stats") {
    req.query = Object.fromEntries(parsedUrl.searchParams.entries());
    res.json = (data) => {
      res.setHeader("Content-Type", "application/json");
      res.end(JSON.stringify(data));
    };
    res.status = (code) => {
      res.statusCode = code;
      return res;
    };
    return statsHandler(req, res);
  }

  // Static HTML route
  if (pathname === "/" || pathname === "/index.html" || pathname === "/dashboard") {
    const filePath = path.join(__dirname, "public", "index.html");
    fs.readFile(filePath, "utf8", (err, data) => {
      if (err) {
        res.statusCode = 500;
        return res.end("Error loading dashboard HTML");
      }
      res.setHeader("Content-Type", "text/html; charset=utf-8");
      res.end(data);
    });
    return;
  }

  res.statusCode = 404;
  res.end("Not found");
});

server.listen(PORT, () => {
  console.log(`🌐 Dashboard running at: ${HOST_BASE}`);
});

// 2. Telegram Polling
async function poll(offset = 0) {
  try {
    const res = await fetch(
      `https://api.telegram.org/bot${BOT_TOKEN}/getUpdates?offset=${offset}&timeout=30`,
    );
    const data = await res.json();
    if (data.ok && data.result.length > 0) {
      for (const update of data.result) {
        if (update.message) {
          console.log(`[${update.message.chat.id}] ${update.message.text || ""}`);
          await handleMessage(update.message, HOST_BASE).catch((e) =>
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

console.log("CatatAI bot running in polling mode...");
fetch(`https://api.telegram.org/bot${BOT_TOKEN}/deleteWebhook`)
  .then(() => poll())
  .catch((e) => {
    console.error("Failed to start polling:", e.message);
    process.exit(1);
  });
