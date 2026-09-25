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

const NEWS_CONFIG = {
  els: {
    terms: ["레이저", "외부 광원", "ELS", "external laser source", "laser"],
    tags: ["CPO", "AI", "레이저", "ELS", "External Laser Source"]
  },
  pmFiber: {
    terms: ["PM fiber", "편광 유지 광섬유", "polarization maintaining fiber", "polarization-maintaining fiber"],
    tags: ["CPO", "AI", "PM fiber", "Polarization Maintaining Fiber"]
  },
  opticalEngine: {
    terms: ["광엔진", "optical engine", "SiPh", "실리콘 포토닉스", "silicon photonics", "PIC", "EIC"],
    tags: ["CPO", "AI", "광엔진", "SiPh", "PIC", "EIC"]
  },
  switchAsic: {
    terms: ["Switch ASIC", "스위치 ASIC", "network switch", "네트워크 스위치"],
    tags: ["CPO", "AI", "Switch ASIC", "Network Switch"]
  },
  package: {
    terms: ["공통 패키지", "advanced packaging", "첨단 패키징", "substrate", "기판", "photonic packaging"],
    tags: ["CPO", "AI", "공통 패키지", "Advanced Packaging", "Substrate"]
  },
  fau: {
    terms: ["FAU", "fiber array unit", "광섬유 어레이", "fiber attach", "fiber-to-chip"],
    tags: ["CPO", "AI", "FAU", "Fiber Array Unit", "Fiber-to-Chip"]
  },
  dataFiber: {
    terms: ["Optical fiber", "광섬유", "fiber cable", "광케이블", "optical connectivity", "AI data center"],
    tags: ["CPO", "AI", "Optical fiber", "AI Data Center", "Optical Connectivity"]
  },
  connector: {
    terms: ["connector", "광커넥터", "VSFF", "MPO", "SN", "MMC", "optical connector"],
    tags: ["CPO", "AI", "Connector", "VSFF", "MPO", "SN", "MMC"]
  }
};

const newsCache = new Map();
const NEWS_CACHE_MS = 15 * 60 * 1000;

function xmlDecode(value = "") {
  return value
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .trim();
}

function tagValue(block, tag) {
  const match = block.match(new RegExp("<" + tag + "(?:\\s[^>]*)?>([\\s\\S]*?)<\\/" + tag + ">", "i"));
  return match ? xmlDecode(match[1]) : "";
}

function buildNewsQuery(company, part) {
  const cfg = NEWS_CONFIG[part] || { terms: ["CPO", "co-packaged optics"], tags: ["CPO", "AI"] };
  const partGroup = cfg.terms.map(term => term.includes(" ") ? '"' + term + '"' : term).join(" OR ");
  return '"' + company + '" (CPO OR "co-packaged optics") (AI OR "AI data center" OR "AI 인프라") (' + partGroup + ') when:30d';
}

function parseGoogleNewsRss(xml) {
  const cutoff = Date.now() - 30 * 24 * 60 * 60 * 1000;
  const blocks = xml.match(/<item>[\s\S]*?<\/item>/gi) || [];
  return blocks.map(block => {
    const source = tagValue(block, "source");
    let title = tagValue(block, "title");
    if (source && title.endsWith(" - " + source)) title = title.slice(0, -(source.length + 3));
    const link = tagValue(block, "link");
    const pubDate = tagValue(block, "pubDate");
    const ts = Date.parse(pubDate);
    return {
      title,
      url: link,
      source,
      publishedAt: Number.isFinite(ts) ? new Date(ts).toISOString() : null,
      date: Number.isFinite(ts) ? new Date(ts).toISOString().slice(0, 10) : ""
    };
  }).filter(item =>
    item.title &&
    item.url &&
    item.publishedAt &&
    Date.parse(item.publishedAt) >= cutoff
  ).slice(0, 5);
}

async function fetchRecentNews(company, part) {
  const key = company + "|" + part;
  const cached = newsCache.get(key);
  if (cached && Date.now() - cached.time < NEWS_CACHE_MS) return cached.data;

  const cfg = NEWS_CONFIG[part] || { tags: ["CPO", "AI"] };
  const query = buildNewsQuery(company, part);
  const url = "https://news.google.com/rss/search?q=" + encodeURIComponent(query) +
    "&hl=ko&gl=KR&ceid=KR:ko";

  const response = await fetch(url, {
    headers: {
      "User-Agent": "Mozilla/5.0 CPO-Component-Explorer/1.0",
      "Accept": "application/rss+xml, application/xml, text/xml"
    },
    signal: AbortSignal.timeout(10000)
  });

  if (!response.ok) throw new Error("Google News RSS HTTP " + response.status);
  const xml = await response.text();
  const data = {
    company,
    part,
    query,
    tags: ["#" + company, ...cfg.tags.map(tag => "#" + tag)],
    windowDays: 30,
    items: parseGoogleNewsRss(xml)
  };

  newsCache.set(key, { time: Date.now(), data });
  return data;
}


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
  if (req.url?.startsWith("/api/news")) {
    try {
      const requestUrl = new URL(req.url, "http://localhost");
      const company = (requestUrl.searchParams.get("company") || "").trim().slice(0, 100);
      const part = (requestUrl.searchParams.get("part") || "").trim();
      if (!company || !NEWS_CONFIG[part]) {
        return sendJson(res, 400, { ok: false, error: "company and valid part are required" });
      }
      fetchRecentNews(company, part)
        .then(data => sendJson(res, 200, { ok: true, ...data }))
        .catch(error => sendJson(res, 502, {
          ok: false,
          error: "Recent news lookup failed",
          detail: String(error && error.message ? error.message : error)
        }));
      return;
    } catch (error) {
      return sendJson(res, 400, { ok: false, error: "Bad news request" });
    }
  }

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
