/**
 * Rez Bounce Co: website <-> Google Sheet connection   (v2: all-day prices)
 *
 * Works with the "Rez Bounce Co Bookings" Sheet (tabs: Requests, Rentals, Fees, Travel, Availability).
 *
 * FIRST-TIME SETUP
 *   1. Open the Sheet → Extensions → Apps Script.
 *   2. Delete what's there, paste this whole file, click Save.
 *   3. Deploy → New deployment → gear icon → Web app.
 *        Execute as: Me      Who has access: Anyone
 *      Click Deploy, approve the permissions (Advanced → Go to project), copy the Web app URL.
 *   4. Put that URL in the website's index.html (API_URL near the top of the script).
 *
 * v2 ONE-TIME STEP (adds all-day prices from the owner's 2026 Pricing sheet)
 *   After pasting this file and saving, pick "upgradeToV2" in the function menu at the top
 *   and click Run once. It fills in his real prices, adds the All-day column, and updates the
 *   estimate formulas on the Requests tab. Then deploy a new version (below).
 *
 * UPDATING LATER
 *   Paste the new file → Save → Deploy → Manage deployments → pencil (Edit) →
 *   Version: New version → Deploy.  (The URL stays the same.)
 *
 * WHAT THE WEBSITE CAN SEE
 *   Prices, fees, the travel chart, and which rentals are taken on which dates.
 *   It never gets anyone's name, phone, email or address.
 */

// Who gets an email for each new request. Leave "" to turn emails off.
const NOTIFY_EMAIL = "kewaeats@gmail.com";
// Also email the customer a "we got your request" note? (Sent from the account that owns this script.)
const EMAIL_CUSTOMER = false;
// Where travel distance is measured from.
const BASE_LOCATION = "Jemez Pueblo, NM";

// Sheet layout (matches the demo Sheet)
const FIRST_ROW = 5;          // first request row on Requests
const COL = { submitted: 1, status: 2, name: 3, phone: 4, email: 5, contactVia: 6, social: 7,
              eventType: 8, date: 9, start: 10, end: 11, address: 12, miles: 13,
              firstItem: 14, generator: 18, firstCalc: 19, lastCalc: 29, notes: 30 };
const HOLDS_DATE = ["pending", "confirmed"];   // statuses that make a rental unavailable

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function readRentals_() {
  const sh = ss_().getSheetByName("Rentals");
  const rows = sh.getRange(4, 1, 4, 7).getValues();
  return rows.filter(r => r[0]).map(r => ({
    name: String(r[0]), type: String(r[1]), minHours: Number(r[2]) || 0,
    price: Number(r[3]) || 0, extraHour: Number(r[4]) || 0, units: Number(r[5]) || 1,
    allDay: r[6] === "" || isNaN(Number(r[6])) ? null : Number(r[6])
  }));
}

function readFees_() {
  const v = ss_().getSheetByName("Fees").getRange(3, 2, 7, 1).getValues().map(r => r[0]);
  return { operator: Number(v[0]) || 0, setupBounce: Number(v[1]) || 0, setupBull: Number(v[2]) || 0,
           setupPackage: Number(v[3]) || 0, generator: Number(v[4]) || 0, base: String(v[5] || BASE_LOCATION),
           allDayHours: Number(v[6]) || 8 };
}

function readTravel_() {
  const rows = ss_().getSheetByName("Travel").getRange(4, 1, 20, 4).getValues();
  return rows.filter(r => r[0] !== "" && r[3] !== "").map(r => ({
    from: Number(r[0]) || 0, to: r[1] === "" ? null : Number(r[1]), label: String(r[2]), fee: Number(r[3]) || 0
  }));
}

/** Every upcoming booking that holds a rental: date + item only (no personal info). */
function readBooked_(rentals) {
  const sh = ss_().getSheetByName("Requests");
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return [];
  const n = last - FIRST_ROW + 1;
  const vals = sh.getRange(FIRST_ROW, 1, n, COL.generator).getValues();
  const shown = sh.getRange(FIRST_ROW, 1, n, COL.generator).getDisplayValues();
  const tz = ss_().getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const out = [];
  vals.forEach((r, i) => {
    const status = String(r[COL.status - 1]).trim().toLowerCase();
    const d = r[COL.date - 1];
    if (!r[COL.name - 1] || !(d instanceof Date) || HOLDS_DATE.indexOf(status) < 0) return;
    const date = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    if (date < today) return;
    rentals.forEach((it, k) => {
      if (String(r[COL.firstItem - 1 + k]).trim().toLowerCase() === "yes") {
        out.push({ date: date, item: it.name, start: shown[i][COL.start - 1], end: shown[i][COL.end - 1], status: status });
      }
    });
  });
  return out;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "distance") return json_(distance_(p.to));
    const rentals = readRentals_();
    return json_({ ok: true, rentals: rentals, fees: readFees_(), travel: readTravel_(), booked: readBooked_(rentals) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Driving miles (one way) from the base to an address, using Google Maps. */
function distance_(to) {
  if (!to || String(to).trim().length < 3) return { ok: false, error: "No address" };
  try {
    const dir = Maps.newDirectionFinder().setOrigin(BASE_LOCATION).setDestination(String(to))
      .setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();
    if (!dir.routes || !dir.routes.length) return { ok: false, error: "Not found" };
    const meters = dir.routes[0].legs[0].distance.value;
    return { ok: true, miles: Math.round(meters / 1609.344 * 10) / 10, found: dir.routes[0].legs[0].end_address };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let p;
    try { p = JSON.parse(e.postData.contents); } catch (x) { p = (e && e.parameter) || {}; }

    const need = ["name", "phone", "email", "date", "start", "end", "address"];
    const missing = need.filter(k => !p[k]);
    const items = Array.isArray(p.items) ? p.items : String(p.items || "").split("|").filter(String);
    if (missing.length || !items.length) return json_({ ok: false, error: "missing", fields: missing });
    if (!p.terms) return json_({ ok: false, error: "terms" });

    // Double-booking check, done again here in case someone booked a moment ago.
    const rentals = readRentals_();
    const booked = readBooked_(rentals);
    const taken = [];
    items.forEach(name => {
      const it = rentals.filter(r => r.name === name)[0];
      if (!it) return;
      const used = booked.filter(b => b.date === p.date && b.item === name).length;
      if (used >= it.units) taken.push(name);
    });
    if (taken.length) return json_({ ok: false, error: "taken", taken: taken });

    // First empty row (formulas are already filled in on the blank rows).
    const sh = ss_().getSheetByName("Requests");
    const last = Math.max(sh.getLastRow(), FIRST_ROW);
    const names = sh.getRange(FIRST_ROW, COL.name, last - FIRST_ROW + 1, 1).getValues();
    let row = FIRST_ROW;
    while (row - FIRST_ROW < names.length && names[row - FIRST_ROW][0] !== "") row++;
    if (row > FIRST_ROW) {
      const calc = sh.getRange(row, COL.firstCalc);
      if (calc.getFormula() === "") {
        sh.getRange(row - 1, COL.firstCalc, 1, COL.lastCalc - COL.firstCalc + 1)
          .copyTo(sh.getRange(row, COL.firstCalc, 1, COL.lastCalc - COL.firstCalc + 1));
      }
    }

    const yes = rentals.map(r => items.indexOf(r.name) >= 0 ? "Yes" : "No");
    while (yes.length < 4) yes.push("No");
    const values = [[
      new Date(), "Pending", p.name, p.phone, p.email, p.contactVia || "", p.social || "",
      p.eventType || "", p.date, p.start, p.end, p.address,
      p.miles === "" || p.miles == null ? "" : Number(p.miles),
      yes[0], yes[1], yes[2], yes[3], p.generator === "Yes" ? "Yes" : "No"
    ]];
    sh.getRange(row, 1, 1, values[0].length).setValues(values);
    const note = [p.notes || "", p.estimate ? "Website estimate: $" + p.estimate : "", p.milesSource ? "Miles: " + p.milesSource : ""]
      .filter(String).join(" · ");
    sh.getRange(row, COL.notes).setValue(note);

    notify_(p, items, row);
    return json_({ ok: true, row: row });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function notify_(p, items, row) {
  try {
    const link = ss_().getUrl() + "#gid=" + ss_().getSheetByName("Requests").getSheetId() + "&range=A" + row;
    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, "New booking request: " + p.date + " · " + items.join(", "),
        "New request from the Rez Bounce Co website\n\n" +
        "Name: " + p.name + "\nPhone: " + p.phone + "\nEmail: " + p.email +
        "\nContacted on: " + (p.contactVia || "") + " (" + (p.social || "") + ")" +
        "\nEvent: " + (p.eventType || "") + "\nDate: " + p.date + ", " + p.start + " to " + p.end +
        "\nWhere: " + p.address + (p.miles ? " (" + p.miles + " mi)" : "") +
        "\nRentals: " + items.join(", ") + "\nGenerator: " + (p.generator || "No") +
        (p.estimate ? "\nWebsite estimate: $" + p.estimate : "") +
        (p.notes ? "\nNotes: " + p.notes : "") + "\n\nOpen the Sheet: " + link);
    }
    if (EMAIL_CUSTOMER && p.email) {
      MailApp.sendEmail(p.email, "Rez Bounce Co: we got your request",
        "Hi " + p.name + ",\n\nThanks for your booking request for " + p.date + " (" + items.join(", ") + ").\n" +
        "This is not a confirmed booking yet. We'll check availability and send your final quote soon.\n\nRez Bounce Co");
    }
  } catch (err) { /* email problems never block a request */ }
}

/**
 * ONE-TIME: run from the Apps Script editor (function menu → upgradeToV2 → Run).
 * Puts in the owner's 2026 prices, adds an All-day price column, and updates the
 * rental-cost formulas on Requests so long rentals are capped at the all-day price.
 */
function upgradeToV2() {
  const ss = ss_();
  // Rentals: Item, Type, Minimum hours, Price for minimum hours, Each extra hour, How many he owns, All-day price, Source
  const r = ss.getSheetByName("Rentals");
  r.getRange(3, 7, 1, 2).setValues([["All-day price", "Where the number came from"]]);
  r.getRange(3, 1, 1, 8).setFontWeight("bold").setFontColor("#ffffff").setBackground("#141414").setWrap(true);
  const src = "Owner's 2026 Pricing sheet (Nov 2026)";
  r.getRange(4, 3, 4, 3).setValues([[4, 100, 25], [4, 100, 25], [4, 120, 30], [3, 325, 100]]);
  r.getRange(4, 7, 4, 2).setValues([[160, src], [160, src], [200, src], [500, src]]);
  r.getRange(4, 4, 4, 2).setNumberFormat('"$"#,##0.00');
  r.getRange(4, 7, 4, 1).setNumberFormat('"$"#,##0.00').setBackground("#FFF4C2");
  r.setColumnWidth(7, 110); r.setColumnWidth(8, 260);
  r.getRange("A2").setValue("Yellow cells are inputs. Change a price here and the website and every estimate update. All-day price covers up to the all-day hours on the Fees tab; each hour past that adds the extra-hour rate.");

  // Fees: hours in an all-day rental
  const f = ss.getSheetByName("Fees");
  f.getRange(9, 1, 1, 3).setValues([["Hours in an all-day rental", 8, "Owner's 2026 Pricing sheet"]]);
  f.getRange(9, 2).setBackground("#FFF4C2");
  f.getRange(9, 3).setFontColor("#777777");

  // Requests: rental cost = cheaper of (minimum price + extra hours) and (all-day price + hours past all-day)
  const q = ss.getSheetByName("Requests");
  const last = Math.max(q.getLastRow(), 44);
  const cols = ["N", "O", "P", "Q"];
  const out = [];
  for (let row = FIRST_ROW; row <= last; row++) {
    out.push(cols.map((c, k) => {
      const rr = 4 + k;
      const byHour = `Rentals!$D$${rr}+MAX(0,S${row}-Rentals!$C$${rr})*Rentals!$E$${rr}`;
      const byDay = `Rentals!$G$${rr}+MAX(0,S${row}-Fees!$B$9)*Rentals!$E$${rr}`;
      return `=IF(${c}${row}<>"Yes",0,IF(Rentals!$G$${rr}="",${byHour},MIN(${byHour},${byDay})))`;
    }));
  }
  q.getRange(FIRST_ROW, 20, out.length, 4).setFormulas(out);
}
