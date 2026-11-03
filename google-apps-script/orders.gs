/**
 * OBDTech Store: order alerts.
 * Every order sent from the website is saved in this Google Sheet
 * and emailed to ALERT_EMAIL.
 *
 * Setup (once): see google-apps-script/README.md in this repository.
 */

const ALERT_EMAIL = "obdtech360@gmail.com";
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

  // Basic checks against junk submissions.
  const phone = String(d.phone || "").replace(/[\s-]/g, "");
  if (!/^(\+?964|0)?7\d{9}$/.test(phone) || !Array.isArray(d.items) || d.items.length === 0 || d.website) {
    return ContentService.createTextOutput("ignored");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const book = SpreadsheetApp.getActiveSpreadsheet();
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
