# Order alerts (email + Google Sheet)

When a customer taps "Send order via WhatsApp" on the store, the site also sends the order here.
The script saves it as a row in a Google Sheet and emails it to **obdtech360@gmail.com**.
Each order has an ID like `OB-8VUSM`, which also appears in the customer's WhatsApp message,
so the email, the sheet row and the WhatsApp chat can be matched.

## Setup (about 10 minutes, once)

Do this signed in to the Google account **obdtech360@gmail.com**.

1. Open **sheets.new** to create a new Google Sheet. Name it `OBDTech Orders`.
2. In the sheet: **Extensions → Apps Script**.
3. Delete the sample code, paste everything from `orders.gs` (this folder), and click **Save**.
4. In the function menu at the top, choose **testAlert**, then click **Run**.
   Google asks for permission: **Review permissions → choose the account → Advanced →
   Go to project (unsafe) → Allow**. ("Unsafe" only means Google hasn't reviewed your own script.)
   You should get a test email and see an `Orders` tab with one test row. Delete that row.
5. Click **Deploy → New deployment**. Gear icon → **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).
6. Put that URL in `index.html` at `const ORDER_HOOK = "";` (or send it to Claude to do it).

## Notes

- Free Gmail accounts can send about 100 emails a day from scripts, which is plenty.
- If you change the script later, use **Deploy → Manage deployments → Edit → New version**
  so the same URL keeps working.
- Orders that arrive in the sheet but never in WhatsApp are customers who didn't press send:
  call or message them.
