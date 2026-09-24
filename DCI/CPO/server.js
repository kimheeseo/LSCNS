const http = require("http");
const fs = require("fs");
const path = require("path");
const pkg = require("./package.json");

const PORT = Number(process.env.PORT || 3000);
const HOST = "0.0.0.0";
const ROOT = __dirname;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".webp": "image/webp",
  ".ico": "image/x-icon",
  ".txt": "text/plain; charset=utf-8",
  ".md": "text/markdown; charset=utf-8"
};

function sendJson(res, status, body) {
  const payload = JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(payload),
    "Cache-Control": "no-store"
  });
  res.end(payload);
}

function safeFilePath(urlPath) {
  const raw = decodeURIComponent((urlPath || "/").split("?")[0]);
  const relative = raw === "/" ? "index.html" : raw.replace(/^\/+/, "");
  const resolved = path.resolve(ROOT, relative);
  if (!resolved.startsWith(path.resolve(ROOT) + path.sep) && resolved !== path.resolve(ROOT, "index.html")) {
    return null;
  }
  return resolved;
}

const server = http.createServer((req, res) => {
  if (req.url === "/api/health" || req.url?.startsWith("/api/health?")) {
    return sendJson(res, 200, {
      ok: true,
      app: "CPO Component Explorer",
      version: pkg.version
    });
  }

  let filePath;
  try {
    filePath = safeFilePath(req.url);
  } catch {
    return sendJson(res, 400, { ok: false, error: "Bad request" });
  }

  if (!filePath) {
    return sendJson(res, 403, { ok: false, error: "Forbidden" });
  }

  fs.stat(filePath, (err, stat) => {
    if (!err && stat.isDirectory()) {
      filePath = path.join(filePath, "index.html");
    }

    fs.readFile(filePath, (readErr, data) => {
      if (readErr) {
        if (readErr.code === "ENOENT") {
          fs.readFile(path.join(ROOT, "index.html"), (fallbackErr, fallback) => {
            if (fallbackErr) {
              return sendJson(res, 404, { ok: false, error: "Not found" });
            }
            res.writeHead(200, { "Content-Type": MIME[".html"] });
            res.end(fallback);
          });
          return;
        }
        return sendJson(res, 500, { ok: false, error: "Server error" });
      }

      const ext = path.extname(filePath).toLowerCase();
      res.writeHead(200, {
        "Content-Type": MIME[ext] || "application/octet-stream",
        "Cache-Control": ext === ".html" ? "no-cache" : "public, max-age=3600"
      });
      res.end(data);
    });
  });
});

server.listen(PORT, HOST, () => {
  console.log(`CPO Component Explorer v${pkg.version} listening on http://${HOST}:${PORT}`);
});
