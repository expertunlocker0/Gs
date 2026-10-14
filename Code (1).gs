/**
 * Google Sheets backend for Tool Shop (index.html).
 * Apps Script に貼り付けて使うサーバー側コードやで。
 * 初回は SPREADSHEET_ID と ADMIN_KEY を必ず変更してから setup() を実行してな。
 */
const SPREADSHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";
const ADMIN_KEY = "CHANGE_THIS_TO_A_LONG_RANDOM_KEY";

const SHEETS = {
  Products: ["ID", "Name", "Category", "Price", "Stock", "Unit", "Description", "ImageURL", "Active", "CreatedAt"],
  Orders: ["OrderID", "CreatedAt", "Customer", "Phone", "Address", "Note", "ItemsJSON", "Total", "Status", "UpdatedAt"],
  Settings: ["Key", "Value"]
};

const SAMPLE_PRODUCTS = [
  ["PT-0001", "ক্ল অথর হাতুড়ি", "হাতিয়ার", 550, 24, "পিস", "কাঠ আর মেটালের কাজে টেকসই হাতুড়ি।", "", true],
  ["PT-0002", "স্ক্রু ড্রাইভার সেট (১২ পিস)", "হাতিয়ার", 950, 15, "সেট", "ফ্ল্যাট, ফিলিপস আর টর্ক ড্রাইভার একসাথে।", "", true],
  ["PT-0003", "মেজারিং টেপ ৫ মিটার", "মাপার যন্ত্র", 250, 40, "পিস", "লকিং ক্লিপ সহ স্টিল টেপ।", "", true],
  ["PT-0004", "কর্ডলেস ড্রিল মেশিন", "পাওয়ার টুল", 4500, 8, "পিস", "১৮V ব্যাটারি, ভেরিয়েবল স্পিড।", "", true],
  ["PT-0005", "অ্যাডজাস্টেবল রেঞ্চ", "হাতিয়ার", 720, 18, "পিস", "৮ ইঞ্চি, ক্রোম ভ্যানেডিয়াম স্টিল।", "", true],
  ["PT-0006", "অ্যাঙ্গেল গ্রাইন্ডার", "পাওয়ার টুল", 3800, 6, "পিস", "কাটা ও ঘষার জন্য ৭৫০W মেশিন।", "", true],
  ["PT-0007", "কম্বিনেশন প্লায়ার", "হাতিয়ার", 480, 30, "পিস", "ইনসুলেটেড গ্রিপ, তার ধরার জন্য।", "", true],
  ["PT-0008", "স্পিরিট লেভেল", "মাপার যন্ত্র", 650, 12, "পিস", "অ্যালুমিনিয়াম বডি, তিন ভায়াল।", "", true],
  ["PT-0009", "ওয়েল্ডিং মেশিন", "ওয়েল্ডিং", 12500, 3, "পিস", "ইনভার্টার টাইপ, ২০০A।", "", true],
  ["PT-0010", "টুল বক্স ২৪ ইঞ্চি", "স্টোরেজ", 2100, 9, "পিস", "মেটাল টুল বক্স, লক সহ।", "", true]
];

/** একবার Apps Script editor থেকে হাতে চালাবে। শিট ও demo product বানাবে। */
function setup() {
  const ss = spreadsheet_();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = SHEETS[name];
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    else {
      const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
      if (existing.slice(0, headers.length).join("|") !== headers.join("|")) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#22384a").setFontColor("#ffffff");
  });
  const products = ss.getSheetByName("Products");
  if (products.getLastRow() <= 1) {
    const stamp = now_();
    const rows = SAMPLE_PRODUCTS.map(function (p) { return p.concat([stamp]); });
    products.getRange(2, 1, rows.length, SHEETS.Products.length).setValues(rows);
  }
  SpreadsheetApp.flush();
  Logger.log("Ready: " + ss.getUrl());
}

function doGet() {
  return json_({ ok: true, message: "Tool Shop API is running" });
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const req = JSON.parse(raw);
    const action = String(req.action || "");

    // Customer-facing read/order actions need no admin key.
    if (action === "ping") return json_({ ok: true, message: "Connected" });
    if (action === "listProducts") return json_({ ok: true, products: listProducts_() });
    if (action === "placeOrder") return json_(placeOrder_(req));
    if (action === "trackOrder") return json_({ ok: true, order: trackOrder_(String(req.orderId || "").trim()) });
    if (action === "getSettings") return json_({ ok: true, settings: getSettings_() });

    // All other actions change or expose admin data.
    if (!authorized_(req.key)) return json_({ ok: false, error: "ভুল অ্যাডমিন কী" });
    if (action === "listOrders") return json_({ ok: true, orders: listOrders_() });
    if (action === "updateStatus") return json_(updateStatus_(req));
    if (action === "deleteOrder") return json_(deleteOrder_(String(req.orderId || "")));
    if (action === "saveProduct") return json_(saveProduct_(req.product || {}));
    if (action === "deleteProduct") return json_(deleteProduct_(String(req.id || "")));
    if (action === "saveSettings") return json_(saveSettings_(req.settings || {}));

    return json_({ ok: false, error: "Unknown action: " + action });
  } catch (err) {
    return json_({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function spreadsheet_() {
  const id = String(SPREADSHEET_ID || "").trim();
  if (!id || id === "PASTE_YOUR_GOOGLE_SHEET_ID_HERE") {
    throw new Error("Code.gs-এর SPREADSHEET_ID-তে Google Sheet ID বসাও");
  }
  return SpreadsheetApp.openById(id);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function authorized_(key) {
  return String(key || "") === String(ADMIN_KEY);
}

function now_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Dhaka", "yyyy-MM-dd HH:mm:ss");
}

function dateKey_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Dhaka", "yyyyMMdd");
}

function bool_(v) {
  return v === true || String(v).toLowerCase() === "true" || String(v) === "1" || String(v).toLowerCase() === "yes";
}

function records_(sheetName) {
  const sh = spreadsheet_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const headers = SHEETS[sheetName];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  return values.map(function (row, i) {
    const obj = { _row: i + 2 };
    headers.forEach(function (h, j) { obj[h] = row[j]; });
    return obj;
  }).filter(function (obj) { return obj[headers[0]] !== "" && obj[headers[0]] !== null; });
}

function productObject_(r) {
  return {
    id: String(r.ID), name: String(r.Name || ""), category: String(r.Category || "অন্যান্য"),
    price: Number(r.Price) || 0, stock: Number(r.Stock) || 0, unit: String(r.Unit || "পিস"),
    description: String(r.Description || ""), imageUrl: String(r.ImageURL || ""),
    active: bool_(r.Active), createdAt: String(r.CreatedAt || "")
  };
}

function listProducts_() {
  return records_("Products").map(productObject_);
}

function orderObject_(r) {
  let items = [];
  try { items = JSON.parse(String(r.ItemsJSON || "[]")); } catch (e) { items = []; }
  return {
    orderId: String(r.OrderID), createdAt: String(r.CreatedAt || ""), customer: String(r.Customer || ""),
    phone: String(r.Phone || ""), address: String(r.Address || ""), note: String(r.Note || ""),
    items: items, total: Number(r.Total) || 0, status: String(r.Status || "pending"), updatedAt: String(r.UpdatedAt || "")
  };
}

function listOrders_() {
  return records_("Orders").map(orderObject_).reverse();
}

/** অর্ডারের দাম/স্টক Sheet থেকে যাচাই করে, browser পাঠানো দাম বিশ্বাস করে না। */
function placeOrder_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const customer = String(req.customer || "").trim();
    const phone = String(req.phone || "").trim();
    const address = String(req.address || "").trim();
    if (!customer || !phone || !address) throw new Error("নাম, ফোন ও ঠিকানা লাগবে");

    const requested = Array.isArray(req.items) ? req.items : [];
    if (!requested.length) throw new Error("অর্ডারে কোনো আইটেম নেই");
    const qtyById = {};
    requested.forEach(function (it) {
      const id = String(it.id || "");
      const qty = Number(it.qty);
      if (!id || !Number.isInteger(qty) || qty < 1) throw new Error("অর্ডারের আইটেম/পরিমাণ ঠিক নেই");
      qtyById[id] = (qtyById[id] || 0) + qty;
    });

    const sh = spreadsheet_().getSheetByName("Products");
    const rows = records_("Products");
    const byId = {};
    rows.forEach(function (r) { byId[String(r.ID)] = r; });
    const orderItems = [];
    let total = 0;
    Object.keys(qtyById).forEach(function (id) {
      const r = byId[id];
      if (!r || !bool_(r.Active)) throw new Error("একটি প্রোডাক্ট আর পাওয়া যাচ্ছে না");
      const qty = qtyById[id];
      const stock = Number(r.Stock) || 0;
      if (stock < qty) throw new Error(String(r.Name) + "-এর যথেষ্ট স্টক নেই");
      const price = Number(r.Price) || 0;
      orderItems.push({ id: id, name: String(r.Name), price: price, qty: qty });
      total += price * qty;
    });

    // Stock update, still protected by the script lock.
    rows.forEach(function (r) {
      const q = qtyById[String(r.ID)] || 0;
      if (q) sh.getRange(r._row, 5).setValue((Number(r.Stock) || 0) - q);
    });

    const orders = spreadsheet_().getSheetByName("Orders");
    const prefix = "ORD-" + dateKey_() + "-";
    let max = 0;
    records_("Orders").forEach(function (r) {
      const id = String(r.OrderID || "");
      if (id.indexOf(prefix) === 0) max = Math.max(max, Number(id.slice(prefix.length)) || 0);
    });
    const orderId = prefix + String(max + 1).padStart(4, "0");
    const stamp = now_();
    const order = {
      orderId: orderId, createdAt: stamp, customer: customer, phone: phone, address: address,
      note: String(req.note || ""), items: orderItems, total: total, status: "pending", updatedAt: stamp
    };
    orders.appendRow([orderId, stamp, customer, phone, address, order.note, JSON.stringify(orderItems), total, "pending", stamp]);
    SpreadsheetApp.flush();
    return { ok: true, order: order };
  } finally {
    lock.releaseLock();
  }
}

function trackOrder_(id) {
  const found = records_("Orders").filter(function (r) { return String(r.OrderID) === id; })[0];
  if (!found) return null;
  const o = orderObject_(found);
  // Public tracking returns no personal contact/address details.
  return { orderId: o.orderId, createdAt: o.createdAt, items: o.items, total: o.total, status: o.status, updatedAt: o.updatedAt };
}

function updateStatus_(req) {
  const allowed = ["pending", "confirmed", "packed", "shipped", "delivered", "cancelled"];
  const id = String(req.orderId || "");
  const status = String(req.status || "");
  if (allowed.indexOf(status) < 0) throw new Error("এই status অনুমোদিত নয়");
  const r = records_("Orders").filter(function (x) { return String(x.OrderID) === id; })[0];
  if (!r) throw new Error("অর্ডার পাওয়া যায়নি");
  const sh = spreadsheet_().getSheetByName("Orders");
  sh.getRange(r._row, 9).setValue(status);
  sh.getRange(r._row, 10).setValue(now_());
  return { ok: true };
}

function deleteOrder_(id) {
  const r = records_("Orders").filter(function (x) { return String(x.OrderID) === id; })[0];
  if (!r) throw new Error("অর্ডার পাওয়া যায়নি");
  spreadsheet_().getSheetByName("Orders").deleteRow(r._row);
  return { ok: true };
}

function nextProductId_() {
  let max = 0;
  records_("Products").forEach(function (r) {
    const m = String(r.ID || "").match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return "PT-" + String(max + 1).padStart(4, "0");
}

function saveProduct_(p) {
  const sh = spreadsheet_().getSheetByName("Products");
  const id = String(p.id || "").trim() || nextProductId_();
  const existing = records_("Products").filter(function (r) { return String(r.ID) === id; })[0];
  const old = existing ? productObject_(existing) : {};
  const row = [
    id,
    p.name !== undefined ? String(p.name) : (old.name || ""),
    p.category !== undefined ? String(p.category) : (old.category || "অন্যান্য"),
    p.price !== undefined ? Number(p.price) || 0 : (old.price || 0),
    p.stock !== undefined ? Number(p.stock) || 0 : (old.stock || 0),
    p.unit !== undefined ? String(p.unit) : (old.unit || "পিস"),
    p.description !== undefined ? String(p.description) : (old.description || ""),
    p.imageUrl !== undefined ? String(p.imageUrl) : (old.imageUrl || ""),
    p.active !== undefined ? Boolean(p.active) : (existing ? old.active : true),
    existing ? (existing.CreatedAt || now_()) : now_()
  ];
  if (existing) sh.getRange(existing._row, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, product: productObject_({ ID:row[0], Name:row[1], Category:row[2], Price:row[3], Stock:row[4], Unit:row[5], Description:row[6], ImageURL:row[7], Active:row[8], CreatedAt:row[9] }) };
}

function deleteProduct_(id) {
  const r = records_("Products").filter(function (x) { return String(x.ID) === id; })[0];
  if (!r) throw new Error("প্রোডাক্ট পাওয়া যায়নি");
  spreadsheet_().getSheetByName("Products").deleteRow(r._row);
  return { ok: true };
}

function getSettings_() {
  const sh = spreadsheet_().getSheetByName("Settings");
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0] !== "") out[String(r[0])] = r[1];
  });
  return out;
}

function saveSettings_(settings) {
  const sh = spreadsheet_().getSheetByName("Settings");
  const current = getSettings_();
  Object.keys(settings).forEach(function (key) { current[key] = settings[key]; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const rows = Object.keys(current).map(function (key) { return [key, current[key]]; });
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
  return { ok: true };
}
