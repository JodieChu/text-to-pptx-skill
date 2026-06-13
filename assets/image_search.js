// image_search.js — 为 text-to-pptx 自动配图（借鉴 ppt-master 的 image_search 思路）
// 依赖：仅 Node 内置模块 + 系统 curl（无需 npm 包）。
//
// 用法（代码内）：
//   const IS = require("./image_search");
//   const img = await IS.searchImage("人工智能 芯片", { orient:"landscape" });
//   // img = { path, w, h, source, attribution, license, sourceUrl } 或 null（失败时优雅降级）
//   if (img) slide.addImage({ path: img.path, x, y, w, h, sizing:{type:"cover", w, h} });
//
// 用法（命令行自测）：
//   node image_search.js "量子计算" landscape
//
// 图源（按可达性/是否需 key）：
//   - wikimedia（默认，关键词，直连，免 key，CC/公共领域，带署名）   ← 你的网络实测直连可达
//   - pexels / pixabay（更好看的图库照片，需各自 API key，走环境变量）
//   - openverse（聚合 CC 图，需走代理，设 TTP_PROXY 环境变量）
//   - picsum（无关键词的装饰图/兜底，直连）
//   source:"auto"（默认）= wikimedia → picsum 兜底
const { execFileSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const HOME = process.env.HOME || process.env.USERPROFILE || ".";
const DEFAULT_CACHE = path.join(HOME, "ai_projects", ".tools-cache", "img-cache");
const PROXY = process.env.TTP_PROXY || process.env.HTTPS_PROXY || process.env.https_proxy || "";

function curl(url, { proxy = "", timeout = 15, out = "" } = {}) {
  const args = ["-sL", "--max-time", String(timeout)];
  if (proxy) args.push("-x", proxy);
  if (out) { args.push("-o", out, "-w", "%{http_code}"); }
  try {
    const r = execFileSync("curl", [...args, url], { encoding: out ? "utf8" : "utf8", maxBuffer: 64 * 1024 * 1024 });
    return out ? r.trim() : r; // out 模式返回 http_code，否则返回 body
  } catch (e) { return out ? "000" : ""; }
}

function ensureDir(d) { fs.mkdirSync(d, { recursive: true }); }
function cachePath(cacheDir, key, ext) { return path.join(cacheDir, crypto.createHash("sha1").update(key).digest("hex").slice(0, 16) + ext); }

// ---------- 各图源 ----------
function fromWikimedia(query, { orient, minWidth, targetW, proxy }) {
  const q = encodeURIComponent(`filetype:bitmap ${query}`);
  const api = `https://commons.wikimedia.org/w/api.php?action=query&format=json&prop=imageinfo&generator=search&gsrsearch=${q}&gsrnamespace=6&gsrlimit=12&iiprop=url|size|extmetadata&iiurlwidth=${targetW}`;
  let data; try { data = JSON.parse(curl(api, { proxy, timeout: 12 })); } catch { return null; }
  const pages = data && data.query && data.query.pages ? Object.values(data.query.pages) : [];
  const cands = pages.map(p => p.imageinfo && p.imageinfo[0]).filter(Boolean)
    .filter(ii => (ii.width || 0) >= minWidth)
    .filter(ii => {
      if (orient === "landscape") return ii.width >= ii.height;
      if (orient === "portrait") return ii.height >= ii.width;
      return true;
    });
  const pick = cands[0]; if (!pick) return null;
  const meta = pick.extmetadata || {};
  const strip = s => (s ? String(s).replace(/<[^>]+>/g, "").trim() : "");
  return {
    url: pick.thumburl || pick.url, w: pick.thumbwidth || pick.width, h: pick.thumbheight || pick.height,
    source: "Wikimedia Commons",
    attribution: strip(meta.Artist && meta.Artist.value) || "Wikimedia Commons",
    license: strip(meta.LicenseShortName && meta.LicenseShortName.value) || "见来源页",
    sourceUrl: pick.descriptionurl || pick.url,
  };
}

function fromPexels(query, { orient, proxy }) {
  const key = process.env.PEXELS_KEY; if (!key) return null;
  const o = orient === "portrait" ? "&orientation=portrait" : orient === "landscape" ? "&orientation=landscape" : "";
  const api = `https://api.pexels.com/v1/search?query=${encodeURIComponent(query)}&per_page=1${o}`;
  let body; try {
    body = execFileSync("curl", ["-sL", "--max-time", "12", ...(proxy ? ["-x", proxy] : []), "-H", `Authorization: ${key}`, api], { encoding: "utf8" });
  } catch { return null; }
  let d; try { d = JSON.parse(body); } catch { return null; }
  const p = d.photos && d.photos[0]; if (!p) return null;
  return { url: p.src.large2x || p.src.large, w: p.width, h: p.height, source: "Pexels", attribution: p.photographer || "Pexels", license: "Pexels License (免费可商用)", sourceUrl: p.url };
}

function fromPixabay(query, { orient, proxy }) {
  const key = process.env.PIXABAY_KEY; if (!key) return null;
  const o = orient === "portrait" ? "&orientation=vertical" : orient === "landscape" ? "&orientation=horizontal" : "";
  const api = `https://pixabay.com/api/?key=${key}&q=${encodeURIComponent(query)}&image_type=photo&per_page=3${o}`;
  let d; try { d = JSON.parse(curl(api, { proxy, timeout: 12 })); } catch { return null; }
  const h = d.hits && d.hits[0]; if (!h) return null;
  return { url: h.largeImageURL || h.webformatURL, w: h.imageWidth, h: h.imageHeight, source: "Pixabay", attribution: h.user || "Pixabay", license: "Pixabay License (免费可商用)", sourceUrl: h.pageURL };
}

function fromOpenverse(query, { orient, proxy }) {
  const o = orient === "portrait" ? "&aspect_ratio=tall" : orient === "landscape" ? "&aspect_ratio=wide" : "";
  const api = `https://api.openverse.org/v1/images/?q=${encodeURIComponent(query)}&page_size=1&license_type=commercial${o}`;
  // Openverse 你的网络需走代理
  let d; try { d = JSON.parse(curl(api, { proxy: proxy || PROXY || "http://127.0.0.1:7890", timeout: 12 })); } catch { return null; }
  const r = d.results && d.results[0]; if (!r) return null;
  return { url: r.url, w: r.width, h: r.height, source: "Openverse", attribution: r.creator || "Openverse", license: r.license || "CC", sourceUrl: r.foreign_landing_url || r.url };
}

function fromPicsum(query, { orient }) {
  const w = orient === "portrait" ? 800 : 1200, h = orient === "portrait" ? 1100 : 700;
  const seed = encodeURIComponent(query || "deck");
  return { url: `https://picsum.photos/seed/${seed}/${w}/${h}`, w, h, source: "Lorem Picsum", attribution: "Lorem Picsum", license: "免费装饰图（无关键词）", sourceUrl: "https://picsum.photos" };
}

// ---------- 主入口 ----------
async function searchImage(query, opts = {}) {
  const orient = opts.orient || "landscape";
  const minWidth = opts.minWidth || 600;
  const targetW = opts.targetW || (orient === "portrait" ? 1000 : 1400);
  const cacheDir = opts.cacheDir || DEFAULT_CACHE;
  const proxy = opts.proxy !== undefined ? opts.proxy : "";
  ensureDir(cacheDir);

  let order;
  switch (opts.source || "auto") {
    case "wikimedia": order = ["wikimedia"]; break;
    case "pexels": order = ["pexels"]; break;
    case "pixabay": order = ["pixabay"]; break;
    case "openverse": order = ["openverse"]; break;
    case "picsum": order = ["picsum"]; break;
    case "best": order = ["pexels", "pixabay", "openverse", "wikimedia", "picsum"]; break; // 有 key 时优先好图
    default: order = ["wikimedia", "picsum"]; // auto：免 key 直连优先
  }
  const fns = { wikimedia: fromWikimedia, pexels: fromPexels, pixabay: fromPixabay, openverse: fromOpenverse, picsum: fromPicsum };

  for (const name of order) {
    let hit; try { hit = fns[name](query, { orient, minWidth, targetW, proxy }); } catch { hit = null; }
    if (!hit || !hit.url) continue;
    const ext = (hit.url.match(/\.(jpe?g|png|webp|gif)(\?|$)/i) || [, "jpg"])[1].toLowerCase().replace("jpeg", "jpg");
    const file = cachePath(cacheDir, hit.url, "." + ext);
    if (fs.existsSync(file) && fs.statSync(file).size > 2000) return { ...hit, path: file };
    const useProxy = name === "openverse" ? (proxy || PROXY || "http://127.0.0.1:7890") : proxy;
    const code = curl(hit.url, { proxy: useProxy, timeout: 20, out: file });
    if (code === "200" && fs.existsSync(file) && fs.statSync(file).size > 2000) return { ...hit, path: file };
    try { fs.existsSync(file) && fs.unlinkSync(file); } catch {}
  }
  return null; // 全部失败 → 调用方应优雅降级（画占位框）
}

module.exports = { searchImage };

// CLI 自测
if (require.main === module) {
  (async () => {
    const q = process.argv[2] || "artificial intelligence";
    const orient = process.argv[3] || "landscape";
    const src = process.argv[4] || "auto";
    const img = await searchImage(q, { orient, source: src });
    console.log(JSON.stringify(img, null, 2));
  })();
}
