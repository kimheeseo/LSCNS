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

const NEWS_LOCALES = {
  ko: { hl: "ko", gl: "KR", ceid: "KR:ko" },
  en: { hl: "en-US", gl: "US", ceid: "US:en" },
  ja: { hl: "ja", gl: "JP", ceid: "JP:ja" },
  zh: { hl: "zh-CN", gl: "CN", ceid: "CN:zh-Hans" },
  de: { hl: "de", gl: "DE", ceid: "DE:de" }
};


const STOCK_TICKERS = {
  broadcom: { yahoo: "AVGO", label: "Broadcom", currency: "USD" },
  nvidia: { yahoo: "NVDA", label: "NVIDIA", currency: "USD" },
  marvell: { yahoo: "MRVL", label: "Marvell", currency: "USD" },
  coherent: { yahoo: "COHR", label: "Coherent", currency: "USD" },
  lumentum: { yahoo: "LITE", label: "Lumentum", currency: "USD" },
  corning: { yahoo: "GLW", label: "Corning", currency: "USD" },
  intel: { yahoo: "INTC", label: "Intel", currency: "USD" },
  cisco: { yahoo: "CSCO", label: "Cisco", currency: "USD" },
  tsmc: { yahoo: "TSM", label: "TSMC ADR", currency: "USD" },
  ase: { yahoo: "ASX", label: "ASE Technology", currency: "USD" },
  amkor: { yahoo: "AMKR", label: "Amkor", currency: "USD" },
  globalfoundries: { yahoo: "GFS", label: "GlobalFoundries", currency: "USD" },
  umc: { yahoo: "UMC", label: "UMC ADR", currency: "USD" },
  fujikura: { yahoo: "5803.T", label: "Fujikura", currency: "JPY" },
  sumitomo: { yahoo: "5802.T", label: "Sumitomo Electric", currency: "JPY" },
  furukawa: { yahoo: "5801.T", label: "Furukawa Electric", currency: "JPY" },
  hengtong: { yahoo: "600487.SS", label: "Hengtong Optic-Electric", currency: "CNY" },
  yofc: { yahoo: "601869.SS", label: "YOFC", currency: "CNY" },
  ztt: { yahoo: "600522.SS", label: "ZTT", currency: "CNY" },
  lscns: { yahoo: "006260.KS", label: "LS reference", currency: "KRW" }
};

const stockCache = new Map();
const STOCK_CACHE_MS = 15 * 60 * 1000;

async function fetchStockSeries(id) {
  const meta = STOCK_TICKERS[id];
  if (!meta) return null;

  const cached = stockCache.get(id);
  if (cached && Date.now() - cached.time < STOCK_CACHE_MS) return cached.data;

  const url = "https://query1.finance.yahoo.com/v8/finance/chart/" +
    encodeURIComponent(meta.yahoo) +
    "?range=6mo&interval=1d&includePrePost=false&events=div%2Csplits";

  const response = await fetch(url, {
    headers: {
      "User-Agent": "Mozilla/5.0 CPO-Component-Explorer/1.0",
      "Accept": "application/json"
    },
    signal: AbortSignal.timeout(10000)
  });

  if (!response.ok) throw new Error("Yahoo Finance HTTP " + response.status);
  const json = await response.json();
  const result = json?.chart?.result?.[0];
  if (!result) throw new Error(json?.chart?.error?.description || "No stock chart data");

  const timestamps = result.timestamp || [];
  const quote = result.indicators?.quote?.[0] || {};
  const closes = quote.close || [];
  const volumes = quote.volume || [];
  const points = [];

  for (let i = 0; i < timestamps.length; i++) {
    const close = closes[i];
    if (typeof close !== "number" || !Number.isFinite(close)) continue;
    points.push({
      t: timestamps[i] * 1000,
      close,
      volume: typeof volumes[i] === "number" && Number.isFinite(volumes[i]) ? volumes[i] : null
    });
  }

  if (!points.length) throw new Error("No valid price points");

  const first = points[0].close;
  const last = points[points.length - 1].close;
  const previous = points.length > 1 ? points[points.length - 2].close : first;
  const change = last - previous;
  const changePct = previous ? (change / previous) * 100 : 0;
  const rangeChangePct = first ? ((last - first) / first) * 100 : 0;

  const data = {
    id,
    symbol: meta.yahoo,
    label: meta.label,
    currency: result.meta?.currency || meta.currency,
    exchangeName: result.meta?.fullExchangeName || result.meta?.exchangeName || "",
    regularMarketPrice: result.meta?.regularMarketPrice ?? last,
    previousClose: result.meta?.chartPreviousClose ?? previous,
    change,
    changePct,
    rangeChangePct,
    points
  };

  stockCache.set(id, { time: Date.now(), data });
  return data;
}


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

function newsAiTerms(lang = "ko") {
  const map = {
    ko: '(AI OR "AI 데이터센터" OR "AI 인프라")',
    en: '(AI OR "AI data center" OR "AI infrastructure")',
    ja: '(AI OR "AIデータセンター" OR "AIインフラ")',
    zh: '(AI OR "AI 数据中心" OR "AI 基础设施")',
    de: '(AI OR "KI-Rechenzentrum" OR "KI-Infrastruktur")'
  };
  return map[lang] || map.ko;
}

function buildStrictNewsQuery(company, part, lang = "ko") {
  const cfg = NEWS_CONFIG[part] || { terms: ["CPO", "co-packaged optics"] };
  const partGroup = cfg.terms.map(term => term.includes(" ") ? '"' + term + '"' : term).join(" OR ");
  return '"' + company + '" (CPO OR "co-packaged optics") ' + newsAiTerms(lang) + ' (' + partGroup + ') when:30d';
}

function buildCpoAiNewsQuery(company, lang = "ko") {
  return '"' + company + '" ((CPO OR "co-packaged optics") OR ' + newsAiTerms(lang) + ') when:30d';
}

function buildCompanyNewsQuery(company) {
  return '"' + company + '" when:30d';
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
  );
}

async function fetchGoogleNewsQuery(query, locale) {
  const url = "https://news.google.com/rss/search?q=" + encodeURIComponent(query) +
    "&hl=" + encodeURIComponent(locale.hl) +
    "&gl=" + encodeURIComponent(locale.gl) +
    "&ceid=" + encodeURIComponent(locale.ceid);

  const response = await fetch(url, {
    headers: {
      "User-Agent": "Mozilla/5.0 CPO-Component-Explorer/1.0",
      "Accept": "application/rss+xml, application/xml, text/xml"
    },
    signal: AbortSignal.timeout(10000)
  });

  if (!response.ok) throw new Error("Google News RSS HTTP " + response.status);
  return parseGoogleNewsRss(await response.text());
}

function normalizeNewsTitle(title = "") {
  return title
    .toLowerCase()
    .replace(/[\s\u00a0]+/g, " ")
    .replace(/[“”"'‘’´.,:;!?()[\]{}<>·•\-–—_]/g, "")
    .trim();
}

function mergeNewsItems(target, incoming, relevance, limit = 5) {
  for (const item of incoming) {
    if (target.length >= limit) break;
    const key = normalizeNewsTitle(item.title);
    const duplicate = target.some(existing =>
      existing.url === item.url ||
      normalizeNewsTitle(existing.title) === key
    );
    if (!duplicate) target.push({ ...item, relevance });
  }
  return target;
}

async function fetchRecentNews(company, part, lang = "ko") {
  const key = company + "|" + part + "|" + lang;
  const cached = newsCache.get(key);
  if (cached && Date.now() - cached.time < NEWS_CACHE_MS) return cached.data;

  const cfg = NEWS_CONFIG[part] || { tags: ["CPO", "AI"] };
  const locale = NEWS_LOCALES[lang] || NEWS_LOCALES.ko;
  const strictQuery = buildStrictNewsQuery(company, part, lang);
  const cpoAiQuery = buildCpoAiNewsQuery(company, lang);
  const companyQuery = buildCompanyNewsQuery(company);

  const items = [];
  const queryLog = [];

  try {
    const strict = await fetchGoogleNewsQuery(strictQuery, locale);
    mergeNewsItems(items, strict, "component", 5);
    queryLog.push({ tier: "component", query: strictQuery, found: strict.length });
  } catch (error) {
    queryLog.push({ tier: "component", query: strictQuery, found: 0, error: String(error.message || error) });
  }

  if (items.length < 5) {
    try {
      const related = await fetchGoogleNewsQuery(cpoAiQuery, locale);
      mergeNewsItems(items, related, "cpo_ai", 5);
      queryLog.push({ tier: "cpo_ai", query: cpoAiQuery, found: related.length });
    } catch (error) {
      queryLog.push({ tier: "cpo_ai", query: cpoAiQuery, found: 0, error: String(error.message || error) });
    }
  }

  if (items.length < 5) {
    try {
      const general = await fetchGoogleNewsQuery(companyQuery, locale);
      mergeNewsItems(items, general, "company", 5);
      queryLog.push({ tier: "company", query: companyQuery, found: general.length });
    } catch (error) {
      queryLog.push({ tier: "company", query: companyQuery, found: 0, error: String(error.message || error) });
    }
  }

  const data = {
    company,
    part,
    lang,
    query: strictQuery,
    queries: queryLog,
    tags: ["#" + company, ...cfg.tags.map(tag => "#" + tag)],
    windowDays: 30,
    requestedCount: 5,
    items: items.slice(0, 5)
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
  if (req.url?.startsWith("/api/stock")) {
    try {
      const requestUrl = new URL(req.url, "http://localhost");
      const id = (requestUrl.searchParams.get("id") || "").trim();
      if (!STOCK_TICKERS[id]) {
        return sendJson(res, 400, { ok: false, error: "valid stock id is required" });
      }
      fetchStockSeries(id)
        .then(data => sendJson(res, 200, { ok: true, ...data }))
        .catch(error => sendJson(res, 502, {
          ok: false,
          error: "Stock lookup failed",
          detail: String(error && error.message ? error.message : error)
        }));
      return;
    } catch (error) {
      return sendJson(res, 400, { ok: false, error: "Bad stock request" });
    }
  }

  if (req.url?.startsWith("/api/news")) {
    try {
      const requestUrl = new URL(req.url, "http://localhost");
      const company = (requestUrl.searchParams.get("company") || "").trim().slice(0, 100);
      const part = (requestUrl.searchParams.get("part") || "").trim();
      const lang = (requestUrl.searchParams.get("lang") || "ko").trim();
      if (!company || !NEWS_CONFIG[part] || !NEWS_LOCALES[lang]) {
        return sendJson(res, 400, { ok: false, error: "company and valid part are required" });
      }
      fetchRecentNews(company, part, lang)
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
