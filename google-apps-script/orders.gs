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
                      "السعر مع التوصيل (د.ع)", "ملاحظات", "وقت التأكيد", "رابط المحادثة", "تم الشحن ✓"];

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
      [35, 130, 170, 150, 120, 240, 120, 230, 120, 110, 80].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    }
    // Order number = highest "#" so far (multi-item orders use several rows).
    let n = sh.getLastRow() > 1
      ? Math.max(0, ...sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => Number(r[0]) || 0)) : 0;
    let total = 0;
    const putImage = (row, url, alt) => {
      if (!url) return;
      try {
        const blob = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getBlob();
        const img = SpreadsheetApp.newCellImage()
          .setSourceUrl("data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes()))
          .setAltTextTitle(alt || "").build();
        sh.getRange(row, 2).setValue(img);
      } catch (err) {
        sh.getRange(row, 2).setFormula('=IMAGE("' + String(url).replace(/"/g, "") + '")');
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
             o.address || "", price, o.note || "", o.confirmedAt || "", "", false]
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
      sh.getRange(first, 11).insertCheckboxes();
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
    return { archived: archived, ok: true, added: fresh.length, skipped: Math.max(0, posted - fresh.filter(o => o.thread !== "web").length),
             website: fresh.filter(o => o.thread === "web").length,
             total: total, tab: name, url: book.getUrl() + "#gid=" + sh.getSheetId() };
  } finally {
    lock.releaseLock();
  }
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
