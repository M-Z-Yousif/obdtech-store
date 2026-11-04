/**
 * OBDTech Store: order alerts.
 * Every order sent from the website is saved in this Google Sheet
 * and emailed to ALERT_EMAIL.
 *
 * Setup (once): see google-apps-script/README.md in this repository.
 */

const ALERT_EMAIL = "obdtech360@gmail.com";
const SHEET_ID = "1cAOifkOgJfJ4fiC-Gy31rBU0ZtIVfZPro3REsStSA7Y"; // OBDTech Orders sheet
const SHEET_NAME = "Orders";
const HEADERS = ["Time", "Order ID", "Name", "Phone", "Governorate", "Area",
                 "Items", "Products (IQD)", "Delivery (IQD)", "Total (IQD)",
                 "Source", "Language", "Notes"];

function doPost(e) {
  let d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput("bad request");
  }
  if (d.type === "shipping") return json_(writeShipping_(d));
  if (d.type === "shipping-config") return json_(shippingConfig_(d));

  // Basic checks against junk submissions.
  const phone = String(d.phone || "").replace(/[\s-]/g, "");
  if (!/^(\+?964|0)?7\d{9}$/.test(phone) || !Array.isArray(d.items) || d.items.length === 0 || d.website) {
    return ContentService.createTextOutput("ignored");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const book = SpreadsheetApp.openById(SHEET_ID);
    const sheet = book.getSheetByName(SHEET_NAME) || book.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    }
    const items = d.items.map(i => `${i.name} × ${i.qty} = ${i.total}`).join("\n");
    sheet.appendRow([
      new Date(), String(d.id || ""), String(d.name || ""), "'" + phone,
      String(d.gov || ""), String(d.area || ""), items,
      Number(d.sub) || 0, Number(d.del) || 0, Number(d.total) || 0,
      String(d.source || ""), String(d.lang || ""), String(d.note || "")
    ]);
  } finally {
    lock.releaseLock();
  }

  MailApp.sendEmail({
    to: ALERT_EMAIL,
    subject: `طلب جديد OBDTech ${d.id || ""} — ${d.name || ""} — ${d.totalText || ""}`,
    body: String(d.text || "") + "\n\nواتساب الزبون: https://wa.me/" + phone.replace(/^0/, "964").replace(/^\+/, "")
  });

  return ContentService.createTextOutput("ok");
}

/** Run this once from the editor to grant permissions and send yourself a test email. */
function testAlert() {
  doPost({ postData: { contents: JSON.stringify({
    id: "OB-TEST", name: "Test", phone: "07801234567", gov: "النجف", area: "test",
    items: [{ name: "ANENG 681", qty: 1, total: "30,000" }], sub: 30000, del: 5000, total: 35000,
    totalText: "35,000 د.ع", source: "test", lang: "ar", note: "",
    text: "طلب تجريبي من موقع OBDTech"
  }) } });
}

/* ------------------------------------------------------------------ *
 * Daily shipping sheet: confirmed Messenger/Instagram orders ("تم الحجز").
 * Claude posts {type:"shipping", key, day, orders:[...]} every night.
 * Each order: {thread, confirmedAt, name, address, phone, model, image, price, note, link, items?:[{model,image}]}
 * A new tab is made per day; orders already sent on an earlier day are skipped.
 * ------------------------------------------------------------------ */
const SHIP_HEADERS = ["#", "صورة الموديل", "الموديل", "الاسم", "رقم الهاتف", "العنوان",
                      "السعر مع التوصيل (د.ع)", "ملاحظات", "وقت التأكيد", "رابط المحادثة", "الحالة"];

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function shippingBook_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SHIP_SHEET_ID");
  let book = null;
  if (id) { try { book = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!book) {
    book = SpreadsheetApp.create("OBDTech — طلبات الشحن اليومية");
    props.setProperty("SHIP_SHEET_ID", book.getId());
    book.getSheets()[0].setName("_sent").appendRow(["key", "day"]);
  }
  if (!book.getSheetByName("_sent")) book.insertSheet("_sent").appendRow(["key", "day"]);
  if (!book.getSheetByName("اقرأني")) {
    const info = book.insertSheet("اقرأني");
    info.setRightToLeft(true);
    info.getRange("A1:A3").setValues([["طلبات الشحن اليومية — OBDTech"],
      ["كل ليلة تُضاف ورقة جديدة باسم التاريخ فيها الطلبات المؤكدة (تم الحجز) من ماسنجر وانستغرام."],
      ["الطلب لا يتكرر في يومين مختلفين."]]);
    info.getRange("A1").setFontWeight("bold").setFontSize(14);
    info.setColumnWidth(1, 600);
  }
  const done = book.getSheetByName("_sent");
  if (!done.isSheetHidden()) done.hideSheet();
  return book;
}

function writeShipping_(d) {
  const shipKey = PropertiesService.getScriptProperties().getProperty("SHIP_KEY");
  if (!shipKey || d.key !== shipKey) return { ok: false, error: "key" };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const book = shippingBook_();
    const sent = book.getSheetByName("_sent");
    const seen = new Set(sent.getLastRow() > 1
      ? sent.getRange(2, 1, sent.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []);
    const posted = (d.orders || []).length;
    const fresh = (d.orders || []).filter(o => {
      const k = String(o.thread) + "|" + String(o.confirmedAt || "");
      return o.thread && (d.sample || !seen.has(k));
    });

    // Website orders from the last 26 hours that were not sent before (not for sample tabs).
    if (!d.sample) try {
      const web = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
      if (web && web.getLastRow() > 1) {
        const since = Date.now() - 26 * 3600 * 1000;
        web.getRange(2, 1, web.getLastRow() - 1, HEADERS.length).getValues().forEach(r => {
          const t = r[0] instanceof Date ? r[0].getTime() : 0;
          const id = String(r[1] || "");
          if (!id || t < since || seen.has("web|" + id)) return;
          fresh.push({
            thread: "web", key: "web|" + id,
            confirmedAt: Utilities.formatDate(r[0], "Asia/Seoul", "yyyy-MM-dd HH:mm"),
            name: String(r[2] || ""), address: [r[4], r[5]].filter(String).join(" - "),
            phone: String(r[3] || "").replace(/^'/, ""), model: String(r[6] || ""), image: "",
            price: Number(r[9]) || "",
            note: "طلب من الموقع " + id + " — تأكد عبر الواتساب قبل الشحن" + (r[12] ? " • " + r[12] : ""),
            link: "https://wa.me/" + String(r[3] || "").replace(/^'/, "").replace(/^0/, "964")
          });
        });
      }
    } catch (err) {}

    const name = String(d.day || Utilities.formatDate(new Date(), "Asia/Baghdad", "yyyy-MM-dd"));
    let sh = book.getSheetByName(name);
    if (!sh) {
      sh = book.insertSheet(name, 0);
      sh.setRightToLeft(true);
      sh.appendRow(SHIP_HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, SHIP_HEADERS.length).setFontWeight("bold").setBackground("#d32f2f").setFontColor("#ffffff");
      [35, 130, 170, 150, 120, 240, 120, 230, 120, 110, 130].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    }
    // Order number = highest "#" so far (multi-item orders use several rows).
    let n = sh.getLastRow() > 1
      ? Math.max(0, ...sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => Number(r[0]) || 0)) : 0;
    let total = 0;
    const putImage = (row, url, alt) => {
      if (!url) return;
      try {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const blob = res.getBlob();
        if (res.getResponseCode() !== 200 || !/^image\//.test(String(blob.getContentType()))) {
          sh.getRange(row, 2).setValue("لا توجد صورة").setFontColor("#999999");
          return;
        }
        const img = SpreadsheetApp.newCellImage()
          .setSourceUrl("data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes()))
          .setAltTextTitle(alt || "").build();
        sh.getRange(row, 2).setValue(img);
      } catch (err) {
        sh.getRange(row, 2).setValue("لا توجد صورة").setFontColor("#999999");
      }
    };
    fresh.forEach(o => {
      n++;
      const price = Number(o.price) || "";
      if (price) total += price;
      // items: [{model, image}] -> one photo row per product; otherwise one row.
      const items = (Array.isArray(o.items) && o.items.length)
        ? o.items : [{ model: o.model || "", image: o.image || "" }];
      const first = sh.getLastRow() + 1;
      items.forEach((it, i) => {
        sh.appendRow(i === 0
          ? [n, "", it.model || "", o.name || "", o.phone ? "'" + o.phone : "",
             o.address || "", price, o.note || "", o.confirmedAt || "", "", STATUS.wait]
          : ["", "", it.model || "", "", "", "", "", "", "", "", ""]);
        const row = sh.getLastRow();
        sh.setRowHeight(row, 120);
        putImage(row, it.image, it.model);
      });
      const rows = items.length;
      if (rows > 1) {
        [1, 4, 5, 6, 7, 8, 9, 10, 11].forEach(c => sh.getRange(first, c, rows, 1).merge());
        sh.getRange(first, 2, rows, 2).setBorder(null, null, null, null, null, true, "#cccccc", SpreadsheetApp.BorderStyle.DOTTED);
      }
      if (o.link) sh.getRange(first, 10).setFormula('=HYPERLINK("' + String(o.link).replace(/"/g, "") + '","' + (o.thread === "web" ? "واتساب الزبون" : "فتح المحادثة") + '")');
      sh.getRange(first, 11).setDataValidation(statusRule_());
      if (/ناقص|يرجى|صوتية|غير صحيح/.test(String(o.note || ""))) sh.getRange(first, 8).setBackground("#fff3cd").setFontWeight("bold");
      sh.getRange(first + rows - 1, 1, 1, SHIP_HEADERS.length)
        .setBorder(null, null, true, null, null, null, "#999999", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      if (!d.sample) sent.appendRow([o.key || (String(o.thread) + "|" + String(o.confirmedAt || "")), name]);
    });
    sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), SHIP_HEADERS.length)
      .setVerticalAlignment("middle").setWrap(true).setFontSize(11);
    sh.getRange("A2:A").setHorizontalAlignment("center");
    sh.getRange("G2:G").setFontWeight("bold");
    sh.getRange("G2:G").setNumberFormat("#,##0");
    sh.getRange("M1").setValue("المجموع (د.ع)").setFontWeight("bold");
    sh.getRange("N1").setFormula("=SUM(G2:G)").setNumberFormat("#,##0").setFontWeight("bold");
    sh.getRange("M2").setValue("عدد الطلبات").setFontWeight("bold");
    sh.getRange("N2").setFormula("=COUNT(A2:A)").setFontWeight("bold");
    trimSheet_(sh, 14);
    let archived = [];
    if (!d.sample) { try { archived = archiveOldTabs_(book); } catch (err) {} }
    let emailed = 0;
    if (!d.sample && d.notify !== false) {
      try { emailed = notifyNightly_(sh, name, book.getUrl() + "#gid=" + sh.getSheetId()); } catch (err) {}
    }
    return { emailed: emailed, archived: archived, ok: true, added: fresh.length, skipped: Math.max(0, posted - fresh.filter(o => o.thread !== "web").length),
             website: fresh.filter(o => o.thread === "web").length,
             total: total, tab: name, url: book.getUrl() + "#gid=" + sh.getSheetId() };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Nightly email to the shipping team ----------------
 * Addresses live only in Script Properties (TEAM_EMAILS, comma separated),
 * set with {type:"shipping-config", key, teamEmails:"a@x.com,b@y.com"}.
 * ------------------------------------------------------------------ */
/* Review mode (default ON): the nightly email goes only to the owner
 * (REVIEW_EMAIL, default ALERT_EMAIL) with a green "approve" button.
 * One tap on it sends that day's email to the team. Turn review off with
 * {type:"shipping-config", key, review:false} to email the team directly. */
const WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxmupIWSj9D2rT_57nxTb_tEhZkCrz40mVrlQR_cyOMebEox4XuDTlI1gMK-2QVq-DC/exec";
const listEmails_ = s => String(s || "").split(/[,;\s]+/).map(x => x.trim()).filter(x => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));

function shippingConfig_(d) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("SHIP_KEY") || d.key !== props.getProperty("SHIP_KEY")) return { ok: false, error: "key" };
  if (typeof d.teamEmails === "string") props.setProperty("TEAM_EMAILS", listEmails_(d.teamEmails).join(","));
  if (typeof d.reviewEmail === "string") props.setProperty("REVIEW_EMAIL", listEmails_(d.reviewEmail).join(","));
  if (typeof d.review === "boolean") props.setProperty("REVIEW_MODE", d.review ? "on" : "off");
  if (typeof d.resetApproval === "string") props.deleteProperty("APPROVED_" + d.resetApproval);
  let cleaned = null;
  if (d.cleanup) cleaned = cleanupTests_();
  const state = { ok: true, cleaned: cleaned, team: props.getProperty("TEAM_EMAILS") || "", reviewEmail: reviewEmail_(),
                  review: reviewOn_() };
  if (d.test) {
    const book = shippingBook_();
    const sh = book.getSheets().find(s => /^\d{4}-\d{2}-\d{2}$/.test(s.getName())) || book.getSheets()[0];
    state.sent = notifyNightly_(sh, sh.getName(), book.getUrl() + "#gid=" + sh.getSheetId(), true);
  }
  return state;
}

function reviewOn_() { return PropertiesService.getScriptProperties().getProperty("REVIEW_MODE") !== "off"; }
function reviewEmail_() { return PropertiesService.getScriptProperties().getProperty("REVIEW_EMAIL") || ALERT_EMAIL; }

function approveSig_(day) {
  const key = PropertiesService.getScriptProperties().getProperty("SHIP_KEY") || "";
  return Utilities.computeHmacSha256Signature("approve|" + day, key)
    .map(b => ("0" + (b & 255).toString(16)).slice(-2)).join("").slice(0, 24);
}

function notifyNightly_(sh, day, url, isTest) {
  const team = PropertiesService.getScriptProperties().getProperty("TEAM_EMAILS") || "";
  if (!reviewOn_()) return team ? emailTeam_(sh, day, url, team, isTest ? "[تجربة] " : "") : 0;
  const approve = WEBAPP_URL + "?approve=" + encodeURIComponent(day) + "&sig=" + approveSig_(day);
  return emailTeam_(sh, day, url, reviewEmail_(), "[للمراجعة] " + (isTest ? "[تجربة] " : ""), team ? approve : "");
}

/** Owner taps the green button in the review email -> the team gets that day's email (once). */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const page = (msg, ok) => HtmlService.createHtmlOutput(
    "<div dir='rtl' style='font-family:Tahoma,Arial;font-size:22px;padding:40px;text-align:center;color:" +
    (ok ? "#1b5e20" : "#b71c1c") + "'>" + msg + "</div>").setTitle("OBDTech الشحن")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
  if (!p.approve) return page("OBDTech", true);
  const day = String(p.approve);
  if (p.sig !== approveSig_(day)) return page("الرابط غير صحيح", false);
  const props = PropertiesService.getScriptProperties();
  const team = props.getProperty("TEAM_EMAILS");
  if (!team) return page("لا توجد ايميلات للفريق بعد", false);
  if (props.getProperty("APPROVED_" + day)) return page("✓ تم الإرسال للفريق مسبقاً (" + day + ")", true);
  const book = shippingBook_();
  const sh = book.getSheetByName(day);
  if (!sh) return page("ورقة " + day + " غير موجودة", false);
  emailTeam_(sh, day, book.getUrl() + "#gid=" + sh.getSheetId(), team, "");
  props.setProperty("APPROVED_" + day, new Date().toISOString());
  return page("✓ تم إرسال طلبات " + day + " للفريق", true);
}

function emailTeam_(sh, day, url, to, prefix, approveUrl) {
  if (!to) return 0;
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues() : [];
  const orders = [];
  rows.forEach(r => {
    if (r[0] !== "" && r[0] != null) orders.push({ n: r[0], models: [r[2]], name: r[3], address: r[5], price: Number(r[6]) || 0, note: r[7] });
    else if (orders.length && r[2]) orders[orders.length - 1].models.push(r[2]);
  });
  const total = orders.reduce((s, o) => s + o.price, 0);
  const fmt = v => Utilities.formatString("%s", Math.round(v)).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const warn = orders.filter(o => /ناقص|يرجى|صوتية|غير صحيح|تعديل|استبدال|مخفي/.test(String(o.note)));
  let low = [];
  try { low = lowStock_(); } catch (err) {}
  const subject = (prefix || "") + "📦 طلبات الشحن " + day + " — " + orders.length + " طلب — " + fmt(total) + " د.ع";
  const list = orders.map(o =>
    "<tr><td style='padding:6px;border-bottom:1px solid #eee'>" + esc(o.n) + "</td>" +
    "<td style='padding:6px;border-bottom:1px solid #eee'><b>" + esc(o.name) + "</b><br>" + esc(o.address) + "</td>" +
    "<td style='padding:6px;border-bottom:1px solid #eee;direction:ltr;text-align:left'>" + o.models.map(esc).join("<br>") + "</td>" +
    "<td style='padding:6px;border-bottom:1px solid #eee;white-space:nowrap'><b>" + fmt(o.price) + "</b></td></tr>").join("");
  const approveBox = approveUrl
    ? "<div style='background:#e8f5e9;border:1px solid #a5d6a7;padding:12px;border-radius:8px;margin:0 0 14px'>" +
      "<b>للمراجعة:</b> افتح الشيت وتأكد من الطلبات، بعدها اضغط الزر الأخضر حتى يوصل هذا الايميل لفريق الشحن." +
      "<p style='margin:10px 0 0'><a href='" + esc(approveUrl) + "' style='background:#2e7d32;color:#fff;padding:12px 22px;border-radius:8px;text-decoration:none;font-weight:bold;display:inline-block'>✅ موافق — ارسل للفريق</a></p></div>"
    : (prefix && prefix.indexOf("للمراجعة") >= 0
      ? "<p style='background:#eee;padding:8px;border-radius:6px'>للمراجعة فقط — لم تُضف ايميلات الفريق بعد.</p>" : "");
  const html =
    "<div dir='rtl' style='font-family:Tahoma,Arial,sans-serif;font-size:15px;color:#222'>" + approveBox +
    "<h2 style='margin:0 0 8px'>طلبات الشحن — " + esc(day) + "</h2>" +
    "<p style='margin:0 0 12px'>عدد الطلبات: <b>" + orders.length + "</b> &nbsp;•&nbsp; المجموع: <b>" + fmt(total) + " د.ع</b></p>" +
    (warn.length ? "<p style='background:#fff3cd;padding:8px;border-radius:6px;margin:0 0 12px'>⚠️ يحتاج تأكيد قبل الشحن: " +
      warn.map(o => "<b>" + esc(o.name) + "</b> (" + esc(o.note) + ")").join("، ") + "</p>" : "") +
    (orders.length ? "<table style='border-collapse:collapse;width:100%;margin-bottom:14px'>" +
      "<tr style='background:#d32f2f;color:#fff'><th style='padding:6px'>#</th><th style='padding:6px'>الزبون والعنوان</th><th style='padding:6px'>الموديل</th><th style='padding:6px'>السعر</th></tr>" +
      list + "</table>" : "<p>لا توجد طلبات جديدة اليوم.</p>") +
    "<p><a href='" + esc(url) + "' style='background:#d32f2f;color:#fff;padding:12px 22px;border-radius:8px;text-decoration:none;font-weight:bold;display:inline-block'>افتح شيت اليوم (الصور وأرقام الهواتف)</a></p>" +
    (low.length ? "<p style='background:#fdecea;padding:8px;border-radius:6px'>📉 مخزون قليل: " +
      low.map(x => "<b>" + esc(x.model) + "</b> (" + x.qty + ")").join("، ") + "</p>" : "") +
    "<p style='color:#777;font-size:12px'>بعد شحن كل طلب غيّر «الحالة» الى ✅ تم الشحن — والمخزون ينقص تلقائياً. الراجع: ↩️ راجع.</p></div>";
  const plain = "طلبات الشحن " + day + ": " + orders.length + " طلب — المجموع " + fmt(total) + " د.ع\n" + url;
  MailApp.sendEmail({ to: to, subject: subject, body: plain, htmlBody: html, name: "OBDTech الشحن" });
  return to.split(",").length;
}

/* ---------------- Keeping the shipping sheet small ----------------
 * Google Sheets allows ~10 million cells per file. A new tab starts with
 * 1000 rows x 26 columns, so each tab is trimmed to what it uses, and daily
 * tabs older than KEEP_DAYS move to a monthly archive file
 * ("OBDTech — أرشيف الشحن YYYY-MM") in the same Google Drive; its link is
 * added to the "اقرأني" tab.
 * ------------------------------------------------------------------ */
const KEEP_DAYS = 30;

function trimSheet_(sh, keepCols) {
  const spareRows = 2;
  const lastRow = Math.max(sh.getLastRow(), 2);
  const maxRows = sh.getMaxRows();
  if (maxRows > lastRow + spareRows) sh.deleteRows(lastRow + spareRows + 1, maxRows - lastRow - spareRows);
  const cols = Math.max(keepCols || sh.getLastColumn(), 1);
  const maxCols = sh.getMaxColumns();
  if (maxCols > cols) sh.deleteColumns(cols + 1, maxCols - cols);
}

function archiveBook_(month, book) {
  const props = PropertiesService.getScriptProperties();
  const prop = "ARCHIVE_" + month;
  let arch = null;
  const id = props.getProperty(prop);
  if (id) { try { arch = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!arch) {
    arch = SpreadsheetApp.create("OBDTech — أرشيف الشحن " + month);
    props.setProperty(prop, arch.getId());
    props.setProperty(prop + "_BLANK", String(arch.getSheets()[0].getSheetId()));
    // List the archive in the info tab.
    const info = book.getSheetByName("اقرأني");
    if (info) info.appendRow(["أرشيف " + month + ": " + arch.getUrl()]);
  }
  return arch;
}

function archiveOldTabs_(book) {
  const today = new Date(Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd") + "T00:00:00Z");
  const moved = [];
  book.getSheets().forEach(sh => {
    const name = sh.getName();
    const m = name.match(/^(\d{4}-\d{2})-\d{2}$/);
    if (!m) return;
    const age = (today - new Date(name + "T00:00:00Z")) / 86400000;
    if (!(age > KEEP_DAYS)) return;
    const arch = archiveBook_(m[1], book);
    if (arch.getSheetByName(name)) { book.deleteSheet(sh); moved.push(name); return; }
    const copy = sh.copyTo(arch).setName(name);
    copy.setRightToLeft(true);
    // Remove the empty default tab of a new archive file.
    const props = PropertiesService.getScriptProperties();
    const blankId = props.getProperty("ARCHIVE_" + m[1] + "_BLANK");
    if (blankId) {
      const blank = arch.getSheets().find(s => String(s.getSheetId()) === blankId);
      if (blank && arch.getSheets().length > 1) arch.deleteSheet(blank);
      props.deleteProperty("ARCHIVE_" + m[1] + "_BLANK");
    }
    book.deleteSheet(sh);
    moved.push(name);
  });
  return moved;
}

/** Run once from the editor: trims every existing tab (also the hidden "_sent" list). */
function trimAllTabs() {
  const book = shippingBook_();
  book.getSheets().forEach(sh => {
    const n = sh.getName();
    trimSheet_(sh, n === "_sent" ? 2 : (n === "اقرأني" ? 1 : 14));
  });
  Logger.log("trimmed " + book.getSheets().length + " tabs");
}

/** Run once from the editor: creates the shipping sheet and logs its link.
 *  The secret key lives only in Script Properties (SHIP_KEY), never in this repository. */
const SETUP_SHIP_KEY = ""; // paste the key here in the editor only, run setupShipping once, then clear it
function setupShipping() {
  if (SETUP_SHIP_KEY) PropertiesService.getScriptProperties().setProperty("SHIP_KEY", SETUP_SHIP_KEY);
  const book = shippingBook_();
  Logger.log(book.getUrl());
}


/* ======================= Stock (المخزون) =======================
 * Tab "المخزون": one row per product. Current = opening + received + movements.
 * Tab "استلام بضاعة": team adds received stock (date, model, qty).
 * Hidden tab "_حركات": every stock change made by the "الحالة" column in a daily tab.
 *   Status effect per unit: بانتظار/ملغي/راجع = 0, تم الشحن = -1. A change writes only the
 *   difference, so shipped -> returned gives +qty back and repeated edits never double count.
 * The daily tab's model text "SGW Bypass Cable × 5" means quantity 5.
 * ============================================================== */
const STATUS = { wait: "⏳ بانتظار الشحن", shipped: "✅ تم الشحن", returned: "↩️ راجع", cancelled: "❌ ملغي" };
const STOCK_TAB = "المخزون", RECV_TAB = "استلام بضاعة", MOVES_TAB = "_حركات";

function statusRule_() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS.wait, STATUS.shipped, STATUS.returned, STATUS.cancelled], true)
    .setAllowInvalid(false).build();
}
const statusEffect_ = v => String(v || "").indexOf("تم الشحن") >= 0 ? -1 : 0;

const START_PRODUCTS = [
  ["UstaOBD", "اسطه OBD, usta"], ["OBDTech BT80-AR", "BT80"], ["LAUNCH CRB3001", "CRB3001"],
  ["Vgate vLinker MC+", "MC+"], ["Vgate vLinker MS", "vlinker ms"], ["Vgate vLinker FS USB", "vlinker fs, ELM327 USB الاصلي, ELM327 USB اصلي"],
  ["Vgate vLinker BM+", "BM+"], ["Kingbolen Ediag Elite", "ediag"], ["MUCAR BT200 MAX AI", "BT200"],
  ["XTOOL A30M V2", "A30M"], ["XTOOL Anyscan A30X", "A30X"], ["TOPDON RLink J2534", "rlink"],
  ["SM2 PRO J2534", "SM2"], ["KTAG V7.020", "ktag"], ["KESS V2 V5.017", "kess"], ["IProg+ V87", "iprog"],
  ["Tactrix OpenPort 2.0", "openport"], ["AUTOOL AS503", "AS503, oil tester"], ["AUTOOL AS505 Super M3", "AS505, ATF"],
  ["AUTOOL AS507", "AS507"], ["AUTOOL AS502", "AS502"], ["QDB-5A", "QDB"], ["YC600", ""], ["PD60+", "PD60"],
  ["XH300", ""], ["IG80-2", "IG80"], ["CNBJ-620", "CNBJ"], ["TopDiag P100", "P100"], ["TopDiag P60", "P60"],
  ["ANENG RT01", "RT01"], ["FNIRSI 1013D", "1013D"], ["FNIRSI 2C53T", "2C53T"], ["FNIRSI 2C23T", "2C23T"],
  ["FNIRSI DSO510", "DSO510"], ["FNIRSI DSO152", "DSO152"], ["FNIRSI CTG-20", "CTG-20, CTG20"],
  ["ANENG 681", "A681"], ["ANENG 682", "A682"], ["ANENG 683", "A683"], ["ANENG P130", "P130"],
  ["ANENG GN807", "GN807"], ["TASI TA500C", "TA500C"], ["TASI TA636A", "TA636A"], ["ANENG SL106", "SL106"],
  ["CGSULIT TPMS90", "TPMS90"], ["CGSULIT TS01", "TS01"],
  ["SGW Bypass Cable", "SGW"], ["OBD Extension 1m", "extension"], ["OBD Adapter US 2018+", ""],
  ["ELM327 Modified USB Switch", "ELM327 Modified, ELM327 تجاري"], ["MINI VCI J2534", "MINI VCI"],
  ["INPA K+DCAN USB", "INPA, K+DCAN"], ["VCDS HEX V2", "VCDS, HEX V2"], ["GODIAG GT106", "GT106"],
  ["Lonauto", ""], ["Kingbolen WiFi", ""], ["KONNWEI MaxOBD", "KONNWEI"], ["XTOOL AD20 Pro", "AD20"],
  ["ThinkDiag 2", "thinkdiag"], ["ANESOK W600", "W600"], ["C43-M Borescope", "C43-M, C43"],
  ["G40-M Borescope", "G40-M, G40"], ["QBH KS05", "KS05"], ["Wire Break Detector", "كاشف قطع"],
];

/** Run once from the editor: creates the stock tabs (every product = 10) and the edit trigger. */
function setupInventory() {
  const book = shippingBook_();
  let inv = book.getSheetByName(STOCK_TAB);
  if (!inv) {
    inv = book.insertSheet(STOCK_TAB, 0);
    inv.setRightToLeft(true);
    inv.getRange(1, 1, 1, 7).setValues([["الموديل", "الكمية الحالية", "حد التنبيه", "الكمية الافتتاحية", "مستلم", "حركات الطلبات", "أسماء أخرى (للمطابقة)"]])
      .setFontWeight("bold").setBackground("#1565c0").setFontColor("#fff");
    inv.setFrozenRows(1);
    const rows = START_PRODUCTS.map(p => [p[0], "", 3, 10, "", "", p[1]]);
    inv.getRange(2, 1, rows.length, 7).setValues(rows);
    [220, 110, 90, 110, 80, 110, 260].forEach((w, i) => inv.setColumnWidth(i + 1, w));
  }
  let recv = book.getSheetByName(RECV_TAB);
  if (!recv) {
    recv = book.insertSheet(RECV_TAB, 1);
    recv.setRightToLeft(true);
    recv.getRange(1, 1, 1, 4).setValues([["التاريخ", "الموديل", "الكمية المستلمة", "ملاحظة"]])
      .setFontWeight("bold").setBackground("#2e7d32").setFontColor("#fff");
    recv.setFrozenRows(1);
    [110, 220, 120, 260].forEach((w, i) => recv.setColumnWidth(i + 1, w));
    recv.getRange("A2:A500").setNumberFormat("yyyy-mm-dd");
  }
  let mv = book.getSheetByName(MOVES_TAB);
  if (!mv) {
    mv = book.insertSheet(MOVES_TAB);
    mv.appendRow(["الوقت", "الموديل", "الكمية", "الحالة", "الطلب", "المستخدم"]);
  }
  if (!mv.isSheetHidden()) mv.hideSheet();
  refreshStockFormulas_(book);
  // Model dropdown in the receive tab.
  recv.getRange("B2:B500").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(inv.getRange("A2:A"), true).setAllowInvalid(false).build());
  // Installable edit trigger on the shipping sheet (once).
  const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "onShipEdit");
  if (!has) ScriptApp.newTrigger("onShipEdit").forSpreadsheet(book).onEdit().create();
  // Existing daily tabs: checkbox -> status list.
  book.getSheets().forEach(sh => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(sh.getName()) || sh.getLastRow() < 2) return;
    sh.getRange(1, 11).setValue("الحالة");
    const n = sh.getLastRow() - 1, vals = sh.getRange(2, 1, n, 11).getValues();
    vals.forEach((r, i) => {
      if (r[0] === "" || r[0] == null) return;
      const c = sh.getRange(i + 2, 11);
      c.clearDataValidations();
      c.setValue(r[10] === true ? STATUS.shipped : STATUS.wait).setDataValidation(statusRule_());
    });
  });
  Logger.log("Inventory ready: " + book.getUrl());
}

function refreshStockFormulas_(book) {
  const inv = book.getSheetByName(STOCK_TAB);
  const n = Math.max(inv.getLastRow() - 1, 1);
  const f = [];
  for (let r = 2; r < n + 2; r++) {
    f.push([
      `=IF(A${r}="","",D${r}+E${r}+F${r})`, null, null,
      `=IF(A${r}="","",SUMIF('${RECV_TAB}'!B:B,A${r},'${RECV_TAB}'!C:C))`,
      `=IF(A${r}="","",SUMIF('${MOVES_TAB}'!B:B,A${r},'${MOVES_TAB}'!C:C))`]);
  }
  f.forEach((row, i) => {
    const r = i + 2;
    inv.getRange(r, 2).setFormula(row[0]);
    inv.getRange(r, 5).setFormula(row[3]);
    inv.getRange(r, 6).setFormula(row[4]);
  });
  inv.getRange("B2:B").setFontWeight("bold").setHorizontalAlignment("center");
  inv.clearConditionalFormatRules();
  inv.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",$B2<=$C2)').setBackground("#fdecea").setFontColor("#b71c1c")
    .setRanges([inv.getRange("A2:G")]).build()]);
}

function stockIndex_(book) {
  const inv = book.getSheetByName(STOCK_TAB);
  if (!inv || inv.getLastRow() < 2) return [];
  return inv.getRange(2, 1, inv.getLastRow() - 1, 7).getValues().filter(r => r[0]).map(r => ({
    name: String(r[0]),
    keys: [String(r[0])].concat(String(r[6] || "").split(",")).map(k => k.trim().toLowerCase()).filter(k => k.length >= 2)
  }));
}

/** "SGW Bypass Cable × 5" -> {name, qty}. Longest matching name/alias wins; unknown models are added (opening 0). */
function matchModel_(book, text) {
  const raw = String(text || "").trim();
  if (!raw) return null;
  const q = (raw.match(/[×x\*]\s*(\d+)\s*$/i) || [])[1];
  const qty = q ? Number(q) : 1;
  const clean = raw.replace(/[×x\*]\s*\d+\s*$/i, "").trim().toLowerCase();
  let best = null, len = 0;
  stockIndex_(book).forEach(p => p.keys.forEach(k => {
    if ((clean === k || clean.indexOf(k) >= 0) && k.length > len) { best = p.name; len = k.length; }
  }));
  if (!best) {
    const inv = book.getSheetByName(STOCK_TAB);
    best = raw.replace(/[×x\*]\s*\d+\s*$/i, "").trim();
    inv.appendRow([best, "", 3, 0, "", "", "أضيف تلقائياً من طلب"]);
    refreshStockFormulas_(book);
  }
  return { name: best, qty: qty };
}

function onShipEdit(e) {
  try {
    const sh = e.range.getSheet(), name = sh.getName();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(name)) return;
    if (e.range.getColumn() > 11 || e.range.getLastColumn() < 11 || e.range.getLastRow() < 2) return;
    const book = sh.getParent();
    const mv = book.getSheetByName(MOVES_TAB);
    if (!mv) return;
    const last = sh.getLastRow();
    const all = sh.getRange(2, 1, last - 1, 11).getValues();
    const user = (e.user && e.user.getEmail && e.user.getEmail()) || "";
    for (let r = Math.max(2, e.range.getRow()); r <= Math.min(last, e.range.getLastRow()); r++) {
      const row = all[r - 2];
      if (row[0] === "" || row[0] == null) continue;           // only the first row of an order holds the status
      const ref = name + "#" + row[0];
      const effect = statusEffect_(row[10]);
      const models = [row[2]];
      for (let k = r - 1; k < all.length && (all[k][0] === "" || all[k][0] == null) && all[k][2]; k++) models.push(all[k][2]);
      const logged = mv.getLastRow() > 1 ? mv.getRange(2, 1, mv.getLastRow() - 1, 5).getValues().filter(m => m[4] === ref) : [];
      models.forEach(t => {
        const m = matchModel_(book, t);
        if (!m) return;
        const done = logged.filter(x => x[1] === m.name).reduce((s, x) => s + (Number(x[2]) || 0), 0);
        const delta = effect * m.qty - done;
        if (delta) mv.appendRow([new Date(), m.name, delta, String(row[10] || ""), ref, user]);
      });
    }
  } catch (err) { console.error(err); }
}

function lowStock_() {
  const book = shippingBook_();
  const inv = book.getSheetByName(STOCK_TAB);
  if (!inv || inv.getLastRow() < 2) return [];
  return inv.getRange(2, 1, inv.getLastRow() - 1, 3).getValues()
    .filter(r => r[0] && r[1] !== "" && Number(r[1]) <= Number(r[2] || 0))
    .map(r => ({ model: r[0], qty: r[1] }));
}


/** Removes test tabs ("نموذج…", anything with "تجربة") and their stock movements. Real date tabs stay. */
function cleanupTests_() {
  const book = shippingBook_();
  const removed = [];
  const testRefs = [];
  book.getSheets().forEach(sh => {
    const n = sh.getName();
    if (/^نموذج/.test(n) || /تجربة/.test(n)) {
      testRefs.push(n + "#", "2000-01-01#");
      removed.push(n);
      book.deleteSheet(sh);
    }
  });
  const mv = book.getSheetByName(MOVES_TAB);
  let rows = 0;
  if (mv && mv.getLastRow() > 1) {
    const vals = mv.getRange(2, 1, mv.getLastRow() - 1, 6).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const ref = String(vals[i][4] || "");
      if (testRefs.some(t => ref.indexOf(t) === 0) || /^2000-01-01#/.test(ref)) { mv.deleteRow(i + 2); rows++; }
    }
  }
  return { tabs: removed, stockRows: rows };
}
