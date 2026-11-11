#!/usr/bin/env node
/* Builds one small share page per product: p/<id>/index.html
   WhatsApp / Facebook / Instagram read its preview tags (product photo, name, price),
   and a visitor who opens it is sent straight to that product on the store (#v-<id>).
   Run after changing products or prices in index.html:  node tools/build-share-pages.js */
const fs = require("fs"), path = require("path");
const ROOT = path.join(__dirname, ".."), SITE = "https://obdtech.net";
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");

function grab(name, open, close) {
  const start = html.indexOf(`const ${name} = ${open}`);
  if (start < 0) throw new Error(`${name} not found in index.html`);
  let i = html.indexOf(open, start), depth = 0;
  for (let j = i; j < html.length; j++) {
    const c = html[j];
    if (c === '"' || c === "'" || c === "`") { const q = c; j++; while (j < html.length && html[j] !== q) { if (html[j] === "\\") j++; j++; } continue; }
    if (c === open) depth++; else if (c === close && --depth === 0) return Function(`return ${html.slice(i, j + 1)}`)();
  }
  throw new Error(`${name}: unbalanced`);
}
const P = grab("P", "[", "]");
const DETAILS = grab("DETAILS", "{", "}");
const EXT_PHOTO = grab("EXT_PHOTO", "{", "}");
const hp = /const HAS_PHOTO = new Set\((\[[^\]]*\])\)/.exec(html);
const HAS_PHOTO = new Set(hp ? JSON.parse(hp[1]) : []);

const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const fmt = n => Number(n).toLocaleString("en-US");
const abs = u => !u ? "" : /^https?:/.test(u) ? u : `${SITE}/${u.replace(/^\.?\//, "")}`;

const outDir = path.join(ROOT, "p");
fs.rmSync(outDir, { recursive: true, force: true });
let n = 0;
for (const p of P) {
  if (!p.id) continue;
  const img = abs(p.img || (HAS_PHOTO.has(p.id) ? `images/${p.id}.jpg` : (EXT_PHOTO[p.id] || "")));
  const title = `${p.brand} ${p.model} — ${p.name}`;
  const price = `السعر ${fmt(p.price)} دينار + التوصيل 5,000 لكل المحافظات`;
  const lines = (DETAILS[p.id] || p.f || []).slice(0, 3);
  const desc = `${price} · ${lines.join(" · ")}`;
  const url = `${SITE}/p/${p.id}/`, go = `/?src=share#v-${p.id}`;
  const page = `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)} | OBDTech</title>
<meta name="description" content="${esc(desc)}">
<link rel="canonical" href="${url}">
<meta property="og:type" content="product">
<meta property="og:site_name" content="OBDTech">
<meta property="og:locale" content="ar_IQ">
<meta property="og:url" content="${url}">
<meta property="og:title" content="${esc(title)}">
<meta property="og:description" content="${esc(desc)}">
${img ? `<meta property="og:image" content="${esc(img)}">\n<meta property="og:image:alt" content="${esc(p.brand + " " + p.model)}">` : ""}
<meta property="product:price:amount" content="${p.price}">
<meta property="product:price:currency" content="IQD">
<meta name="twitter:card" content="summary_large_image">
<meta name="twitter:title" content="${esc(title)}">
<meta name="twitter:description" content="${esc(desc)}">
${img ? `<meta name="twitter:image" content="${esc(img)}">` : ""}
<script>location.replace(${JSON.stringify(go)});</script>
<style>body{font-family:system-ui,Tahoma,sans-serif;background:#EDF0F3;color:#131820;margin:0;display:grid;place-items:center;min-height:100vh;text-align:center;padding:16px}main{background:#fff;border-radius:14px;padding:20px;max-width:420px}img{max-width:100%;border-radius:10px}a{display:inline-block;margin-top:12px;background:#B41E18;color:#fff;padding:10px 18px;border-radius:10px;text-decoration:none}</style>
</head>
<body>
<main>
${img ? `<img src="${esc(img)}" alt="${esc(p.brand + " " + p.model)}">` : ""}
<h1>${esc(title)}</h1>
<p>${esc(price)}</p>
<a href="${go}">شوف التفاصيل والصور واطلب</a>
</main>
</body>
</html>
`;
  fs.mkdirSync(path.join(outDir, p.id), { recursive: true });
  fs.writeFileSync(path.join(outDir, p.id, "index.html"), page);
  n++;
}
console.log(`share pages: ${n} written to p/`);
