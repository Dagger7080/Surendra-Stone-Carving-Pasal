# Order history in Excel — Google Sheet master log, admin dashboard, one-click download

This site has no server, so nothing survives on its own between visits.
Three things work together to fix that for orders:

1. **A `.xlsx` file on every order email** — already built in, no setup
   needed. Each order attaches a real Excel file with that order's details
   to the email FormSubmit sends you (which lands in Gmail, or wherever
   `FORMSUBMIT_ID`/`STORE_EMAIL` in [pasal-checkout.html](pasal-checkout.html)
   points).
2. **A Google Sheet master log** — one running list of *every* order ever
   placed, in one place.
3. **[admin-dashboard.html](admin-dashboard.html)** — a password-gated page
   (not linked from the storefront) where you log in and see the order list
   and sales totals for today, this month, all time, or any custom date
   range you pick, with a daily or monthly breakdown and a one-click Excel
   download of whatever you're looking at.
   [admin-orders.html](admin-orders.html) still exists too, as a simpler
   one-button "download absolutely everything" page if that's all you need.

Items 2 and 3 need the one-time setup below (about 5 minutes, needs a
Google account).

## 1. Create the Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new
   blank spreadsheet. Name it something like **"Pasal Orders"**.
2. Rename the first tab (bottom-left) to `Orders` — the script below writes
   to a tab with that exact name.

## 2. Add the script

1. In the Sheet, open **Extensions → Apps Script**.
2. Delete whatever is in the editor and paste in this:

```javascript
// Change this to your own long, random passphrase before deploying —
// it's what admin-orders.html asks you for before it will hand back any
// order data. Anyone with this value AND the deployment URL below can read
// every order ever placed, so treat it like a password.
var SECRET_KEY = "CHANGE-ME-TO-A-LONG-RANDOM-PASSPHRASE";

// One row per line item, not one row per order — a 2-item order becomes
// 2 rows that repeat the order-level details and differ only in the item
// columns. Matches the layout of the per-order .xlsx attached to each
// order email (see ORDER_ROW_HEADERS in pasal-checkout.html).
var HEADERS = [
  "S. N", "Order number", "Date", "Full name", "Phone", "Email",
  "District", "Delivery address", "Payment method",
  "Item", "Qty", "Unit Price", "Gross Total", "Delivery Charge", "Net Total"
];

function getOrdersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Orders");
  if (!sheet) sheet = ss.insertSheet("Orders");
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

// Splits a delivery fee evenly across `count` line items, rounded to paisa
// (2 decimals). The last share absorbs whatever the rounding left over, so
// the shares always sum back to exactly `delivery` — e.g. Rs 100 across 3
// items becomes 33.33, 33.33, 33.34, not 33.33 x3 = 99.99. Mirrors
// computeDeliveryShares() in pasal-checkout.html so the Sheet log and the
// per-order .xlsx attachment always agree.
function computeDeliveryShares(delivery, count) {
  if (!count) return [];
  var shares = [];
  var allocated = 0;
  for (var i = 0; i < count - 1; i++) {
    var share = Number((delivery / count).toFixed(2));
    shares.push(share);
    allocated += share;
  }
  shares.push(Number((delivery - allocated).toFixed(2)));
  return shares;
}

function doPost(e) {
  var sheet = getOrdersSheet();
  var p = e.parameter;
  var items = [];
  try { items = JSON.parse(p.itemsJson || "[]"); } catch (err) {}
  var delivery = Number(p.deliveryCharge) || 0;
  var shares = computeDeliveryShares(delivery, items.length);
  var serial = Math.max(0, sheet.getLastRow() - 1);   // rows already in the sheet, before this order's

  items.forEach(function (it, idx) {
    var qty = Number(it.qty) || 0;
    var price = Number(it.price) || 0;
    var gross = qty * price;
    var share = shares[idx];
    serial += 1;
    sheet.appendRow([
      serial,
      p.orderNo || "",
      p.dateStr || "",
      p.name || "",
      p.phone || "",
      p.email || "—",
      p.city || "",
      p.address || "",
      p.payment || "",
      it.name || "",
      qty,
      price,
      gross,
      share,
      gross + share
    ]);
  });

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Used by admin-orders.html to fetch every order row as JSON, guarded by
// SECRET_KEY above. Supports JSONP (a ?callback= param) so it can be read
// from the browser without running into cross-origin restrictions.
function doGet(e) {
  var result;
  if (!e.parameter.key || e.parameter.key !== SECRET_KEY) {
    result = { error: "Not authorized." };
  } else {
    var sheet = getOrdersSheet();
    if (sheet.getLastRow() < 2) {
      result = [];
    } else {
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      result = data.slice(1).map(function (row) {
        var obj = {};
        headers.forEach(function (h, i) { obj[h] = row[i]; });
        return obj;
      });
    }
  }
  var payload = JSON.stringify(result);
  var callback = e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + payload + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}
```

**Note on the S. N column:** it's a running count across the whole sheet, computed
from how many rows already exist when each order comes in. Two customers
checking out at the exact same moment could in rare cases race and get
overlapping numbers — harmless (every row still has its own Order number to
key off), just don't treat S. N as a guaranteed-unique ID on its own.

3. Change `SECRET_KEY` at the top to your own passphrase — anything long and
   hard to guess (e.g. mash the keyboard for 20+ characters). This is the
   only thing standing between the public internet and every customer's
   name, phone number, and address, since the deployment itself has to allow
   "Anyone" access (see step 3) for the order form to be able to write to it.
4. Click the **save icon** (or Ctrl+S). Name the project e.g. "Pasal order
   log".

## 3. Deploy it as a Web App

1. Click **Deploy → New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Fill in:
   - **Execute as:** Me (your Google account)
   - **Who has access:** Anyone
4. Click **Deploy**.
5. The first time, Google will ask you to authorize the script — click
   through **Authorize access → (your account) → Advanced → Go to Pasal
   order log (unsafe) → Allow**. This warning is normal for a script you
   wrote yourself that nobody has reviewed; it's only touching this one
   Sheet.
6. Copy the **Web app URL** it gives you (ends in `/exec`).

## 4. Paste the URL into the site

Open [pasal-checkout.html](pasal-checkout.html), find this line near the top
of the `<script>` block:

```javascript
const SHEET_WEBAPP_URL = "";   // <- paste your Apps Script Web App URL here, see order-log-setup.md
```

paste the URL between the quotes, and save. Every order placed from then on
adds a row to the **Orders** tab.

Then do the same in the other two admin files — open each one, find its own
`SHEET_WEBAPP_URL` line near the top of the `<script>` block, and paste in
the **same** URL:

- [admin-orders.html](admin-orders.html)
- [admin-dashboard.html](admin-dashboard.html)

## 5. The admin dashboard — daily / monthly / date-range reports

Open [admin-dashboard.html](admin-dashboard.html) directly in your browser
(it is **not** linked from any page on the storefront — bookmark it
yourself). Log in with the `SECRET_KEY` passphrase you set in step 2 —
optionally ticking "stay logged in on this device" so you don't have to
re-enter it every visit.

Once logged in:

- **Date range** — Today, This month, All time, or a Custom from/to range.
- **View** — Order list (one row per order, matching the columns in the
  order email), Daily (orders and sales totalled per day), or Monthly
  (same, per month).
- The four summary cards (Orders, Sales, Items sold, Avg order value)
  always reflect whatever date range is currently selected.
- **Download this view (.xlsx)** saves whatever the table is currently
  showing as an Excel file — e.g. switch to "This month" + "Daily" and
  download to get a day-by-day breakdown for the month.
- **↻ Refresh** re-fetches the latest rows from the Sheet (the dashboard
  otherwise works off of whatever it loaded at login).

## 6. Download the whole order history in one click

For just a single "give me everything" Excel file with no login screen or
report options, open [admin-orders.html](admin-orders.html), enter the
`SECRET_KEY` passphrase, and click **Download all orders (.xlsx)**.

You can also just open the Google Sheet itself any time and use
**File → Download → Microsoft Excel (.xlsx)**.

## Notes

- **The Sheet is a convenience log, not the order record of truth.** The
  order email (with its own `.xlsx` attachment) is still what actually
  notifies you of a new order — the Sheet write happens silently in the
  background and, if it ever fails (offline, Google having a bad moment,
  URL not set yet), the order still goes through by email exactly as
  before.
- **Both admin pages are plain, unauthenticated static pages** — anyone
  who has the URL and your `SECRET_KEY` can see or download every order.
  Don't link to either from the storefront, and treat the passphrase like a
  password. If you ever suspect it's leaked, change `SECRET_KEY` in the
  Apps Script (Deploy → Manage deployments → edit → New version) and it
  stops working immediately for anyone using the old one — including
  anything a "stay logged in" checkbox saved in a browser.
- **If you ever edit the script**, redeploy via **Deploy → Manage
  deployments → edit (pencil) → New version**, not "New deployment" — a
  fresh deployment gets a different URL, which would mean updating
  `SHEET_WEBAPP_URL` again in all three files: `pasal-checkout.html`,
  `admin-orders.html`, and `admin-dashboard.html`.
