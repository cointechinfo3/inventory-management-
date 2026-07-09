/*************************************************************
 * AIC INVENTORY APP — Apps Script Backend
 * developed by Gogo Foundation
 * Bind to a Google Sheet > Extensions > Apps Script.
 * Run setupSheets() once, then seedSampleData() for demo.
 * Deploy as Web App to get the URL.
 *************************************************************/

// ---------- SCHEMA -------------------------------------------------------
const SCHEMA = {
  Suppliers:    ['SupplierID','SupplierName','ContactPerson','Phone','Email','Address','City','State'],
  Customers:    ['CustomerID','CustomerName','Phone','Email','Address','City','State'],
  Inventory:    ['ItemID','ItemName','SKU','Category','SupplierID','QuantityInStock','UnitCost','UnitPrice','ReorderLevel','DateAdded'],
  Purchases:    ['PurchaseID','SupplierID','PurchaseDate','Status','TotalAmount','Notes'],
  PurchaseItems:['PurchaseItemID','PurchaseID','ItemID','Quantity','UnitCost','Subtotal'],
  Sales:        ['SaleID','CustomerID','SaleDate','Location','Status','TotalAmount','Notes'],
  SaleItems:    ['SaleItemID','SaleID','ItemID','Quantity','UnitPrice','Subtotal'],
  Receipts:     ['ReceiptID','SaleID','ReceiptDate','AmountReceived','Method','Notes'],
  Payments:     ['PaymentID','PurchaseID','PaymentDate','AmountPaid','Method','Notes'],
  Users:        ['UserID','Name','Email','Role','Status','DateAdded','Password'],
  Settings:     ['Key','Value'],
  AuditLog:     ['Timestamp','UserEmail','Action','Entity','EntityID','Details']
};

const ID_FIELD = {
  Suppliers:'SupplierID', Customers:'CustomerID', Inventory:'ItemID',
  Purchases:'PurchaseID', PurchaseItems:'PurchaseItemID',
  Sales:'SaleID', SaleItems:'SaleItemID',
  Receipts:'ReceiptID', Payments:'PaymentID', Users:'UserID'
};

// ---------- SETUP -------------------------------------------------------
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const headers = SCHEMA[name];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,headers.length)
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
  });
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > Object.keys(SCHEMA).length) ss.deleteSheet(def);
  return 'Schema created.';
}

// ---------- WEB APP ENTRY ------------------------------------------------
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Inventory Management System — Gogo Foundation')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- HELPER: date‑to‑string ---------------------------------------
function formatDate_(val) {
  // Converts a Date object to "yyyy-MM-dd" string, or returns unchanged.
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return val;
}

// ---------- GENERIC SHEET HELPERS ----------------------------------------
function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

function sheetToObjects_(name) {
  const sheet = getSheet_(name);
  const rng = sheet.getDataRange().getValues();
  if (rng.length < 2) return [];
  const headers = rng[0];
  const out = [];
  for (let r = 1; r < rng.length; r++) {
    const row = rng[r];
    if (row.every(function(c){ return c === '' || c === null; })) continue;
    const obj = {};
    headers.forEach(function(h, i) {
      let val = row[i];
      val = formatDate_(val);   // ✅ fix date serialisation
      obj[h] = val;
    });
    obj._row = r + 1;
    out.push(obj);
  }
  return out;
}

function groupBy_(rows, key) {
  const map = {};
  rows.forEach(function(r) { (map[r[key]] = map[r[key]] || []).push(r); });
  return map;
}

function indexBy_(rows, key) {
  const map = {};
  rows.forEach(function(r) { map[r[key]] = r; });
  return map;
}

function getNextId_(name) {
  const idField = ID_FIELD[name];
  const rows = sheetToObjects_(name);
  let max = 0;
  rows.forEach(function(r) {
    const n = parseInt(String(r[idField]).replace(/\D/g,''), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  const prefix = idField.replace('ID','').substring(0,3).toUpperCase();
  return prefix + '-' + String(max + 1).padStart(4, '0');
}

function createRecord_(name, data) {
  const sheet = getSheet_(name);
  const headers = SCHEMA[name];
  const idField = ID_FIELD[name];
  if (idField && !data[idField]) data[idField] = getNextId_(name);
  const row = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
  sheet.appendRow(row);
  return data;
}

function updateRecord_(name, id, data) {
  const idField = ID_FIELD[name];
  const rows = sheetToObjects_(name);
  const target = rows.find(function(r) { return String(r[idField]) === String(id); });
  if (!target) throw new Error('Record not found: ' + id);
  const sheet = getSheet_(name);
  const headers = SCHEMA[name];
  headers.forEach(function(h, i) {
    if (data[h] !== undefined && h !== idField) {
      sheet.getRange(target._row, i + 1).setValue(data[h]);
    }
  });
  return Object.assign({}, target, data);
}

function deleteRecord_(name, id) {
  const idField = ID_FIELD[name];
  const rows = sheetToObjects_(name);
  const target = rows.find(function(r) { return String(r[idField]) === String(id); });
  if (!target) return false;
  getSheet_(name).deleteRow(target._row);
  return true;
}

// ---------- CACHE --------------------------------------------------------
function invalidateDashboardCache_() {
  CacheService.getScriptCache().remove('dashboard_v1');
}

// ---------- AUDIT LOG ----------------------------------------------------
function logAction_(action, entity, entityId, details, userEmailOverride) {
  try {
    const email = userEmailOverride || Session.getActiveUser().getEmail() || 'system';
    createRecord_('AuditLog', {
      Timestamp: new Date().toISOString(),
      UserEmail: email,
      Action: action,
      Entity: entity,
      EntityID: entityId || '',
      Details: details ? JSON.stringify(details) : ''
    });
  } catch (e) {}
}

// ---------- AUTH ---------------------------------------------------------
function loginUser(email, password) {
  const users = sheetToObjects_('Users').filter(function(u) {
    return String(u.Email).toLowerCase() === String(email || '').toLowerCase()
      && u.Status === 'Active';
  });
  if (users.length === 0) throw new Error('Invalid credentials');
  const u = users[0];
  if (String(u.Password || '') !== String(password || '')) {
    throw new Error('Invalid credentials');
  }
  logAction_('login', 'Users', u.UserID, { email: email }, email);
  return {
    UserID: u.UserID,
    Name: u.Name,
    Email: u.Email,
    Role: u.Role
  };
}

// ---------- SUPPLIERS ----------------------------------------------------
function getSuppliers() { return sheetToObjects_('Suppliers'); }
function addSupplier(data) {
  const r = createRecord_('Suppliers', data);
  invalidateDashboardCache_();
  logAction_('create', 'Suppliers', r.SupplierID, { name: r.SupplierName });
  return r;
}
function updateSupplier(id, data) {
  const r = updateRecord_('Suppliers', id, data);
  invalidateDashboardCache_();
  logAction_('update', 'Suppliers', id, data);
  return r;
}
function deleteSupplier(id) {
  const inUse = sheetToObjects_('Inventory').some(function(r){ return String(r.SupplierID) === String(id); })
    || sheetToObjects_('Purchases').some(function(r){ return String(r.SupplierID) === String(id); });
  if (inUse) throw new Error('Supplier is referenced by Inventory or Purchases.');
  const r = deleteRecord_('Suppliers', id);
  invalidateDashboardCache_();
  logAction_('delete', 'Suppliers', id, {});
  return r;
}
function getSuppliersPageData() {
  const suppliers = getSuppliers();
  const totals = {};
  sheetToObjects_('Purchases').forEach(function(p) {
    totals[p.SupplierID] = (totals[p.SupplierID]||0) + Number(p.TotalAmount||0);
  });
  const counts = {};
  sheetToObjects_('Inventory').forEach(function(i) {
    counts[i.SupplierID] = (counts[i.SupplierID]||0) + 1;
  });
  return suppliers.map(function(s) {
    return Object.assign({}, s, {
      TotalPurchased: totals[s.SupplierID] || 0,
      ItemCount: counts[s.SupplierID] || 0
    });
  });
}

// ---------- CUSTOMERS ----------------------------------------------------
function getCustomers() { return sheetToObjects_('Customers'); }
function addCustomer(data) {
  const r = createRecord_('Customers', data);
  invalidateDashboardCache_();
  logAction_('create', 'Customers', r.CustomerID, { name: r.CustomerName });
  return r;
}
function updateCustomer(id, data) {
  const r = updateRecord_('Customers', id, data);
  invalidateDashboardCache_();
  logAction_('update', 'Customers', id, data);
  return r;
}
function deleteCustomer(id) {
  const inUse = sheetToObjects_('Sales').some(function(r){ return String(r.CustomerID) === String(id); });
  if (inUse) throw new Error('Customer is referenced by Sales.');
  const r = deleteRecord_('Customers', id);
  invalidateDashboardCache_();
  logAction_('delete', 'Customers', id, {});
  return r;
}
function getCustomersPageData() {
  const customers = getCustomers();
  const totals = {};
  sheetToObjects_('Sales').forEach(function(s) {
    totals[s.CustomerID] = (totals[s.CustomerID]||0) + Number(s.TotalAmount||0);
  });
  const orderCounts = {};
  sheetToObjects_('Sales').forEach(function(s) {
    orderCounts[s.CustomerID] = (orderCounts[s.CustomerID]||0) + 1;
  });
  return customers.map(function(c) {
    return Object.assign({}, c, {
      TotalSpent: totals[c.CustomerID] || 0,
      OrderCount: orderCounts[c.CustomerID] || 0
    });
  });
}

// ---------- INVENTORY ----------------------------------------------------
function getInventory() {
  const items = sheetToObjects_('Inventory');
  const suppliers = indexBy_(sheetToObjects_('Suppliers'), 'SupplierID');
  return items.map(function(it) {
    const sup = suppliers[it.SupplierID];
    const stock = Number(it.QuantityInStock) || 0;
    const cost = Number(it.UnitCost) || 0;
    const price = Number(it.UnitPrice) || 0;
    const reorder = Number(it.ReorderLevel) || 0;
    return Object.assign({}, it, {
      SupplierName: sup ? sup.SupplierName : '—',
      StockValue: stock * cost,
      PotentialRevenue: stock * price,
      PotentialProfit: stock * (price - cost),
      Margin: price > 0 ? ((price - cost) / price) * 100 : 0,
      IsLowStock: stock <= reorder
    });
  });
}
function getInventoryPageData() {
  return { items: getInventory(), suppliers: getSuppliers() };
}
function addInventoryItem(data) {
  data.DateAdded = data.DateAdded || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  data.QuantityInStock = Number(data.QuantityInStock) || 0;
  data.UnitCost = Number(data.UnitCost) || 0;
  data.UnitPrice = Number(data.UnitPrice) || 0;
  data.ReorderLevel = Number(data.ReorderLevel) || 0;
  const r = createRecord_('Inventory', data);
  invalidateDashboardCache_();
  logAction_('create', 'Inventory', r.ItemID, { name: r.ItemName, sku: r.SKU });
  return r;
}
function updateInventoryItem(id, data) {
  if (data.QuantityInStock !== undefined) data.QuantityInStock = Number(data.QuantityInStock) || 0;
  if (data.UnitCost !== undefined) data.UnitCost = Number(data.UnitCost) || 0;
  if (data.UnitPrice !== undefined) data.UnitPrice = Number(data.UnitPrice) || 0;
  if (data.ReorderLevel !== undefined) data.ReorderLevel = Number(data.ReorderLevel) || 0;
  const r = updateRecord_('Inventory', id, data);
  invalidateDashboardCache_();
  logAction_('update', 'Inventory', id, data);
  return r;
}
function deleteInventoryItem(id) {
  const inUse = sheetToObjects_('PurchaseItems').some(r => String(r.ItemID) === String(id))
    || sheetToObjects_('SaleItems').some(r => String(r.ItemID) === String(id));
  if (inUse) throw new Error('Item is referenced by Purchases or Sales.');
  const r = deleteRecord_('Inventory', id);
  invalidateDashboardCache_();
  logAction_('delete', 'Inventory', id, {});
  return r;
}
function adjustStock(itemId, delta, reason) {
  const idField = ID_FIELD.Inventory;
  const rows = sheetToObjects_('Inventory');
  const item = rows.find(r => String(r[idField]) === String(itemId));
  if (!item) throw new Error('Item not found: ' + itemId);
  const newQty = (Number(item.QuantityInStock) || 0) + Number(delta);
  if (newQty < 0) throw new Error('Adjustment would result in negative stock.');
  const sheet = getSheet_('Inventory');
  sheet.getRange(item._row, SCHEMA.Inventory.indexOf('QuantityInStock') + 1).setValue(newQty);
  invalidateDashboardCache_();
  logAction_('adjust', 'Inventory', itemId, { delta, reason, newQty });
  return { ItemID: itemId, QuantityInStock: newQty };
}

// ---------- PURCHASES ----------------------------------------------------
function getPurchases() {
  const purchases = sheetToObjects_('Purchases');
  const suppliers = indexBy_(sheetToObjects_('Suppliers'), 'SupplierID');
  const invIndex = indexBy_(sheetToObjects_('Inventory'), 'ItemID');
  const itemsByPurchase = groupBy_(sheetToObjects_('PurchaseItems'), 'PurchaseID');
  const paymentsByPurchase = groupBy_(sheetToObjects_('Payments'), 'PurchaseID');
  return purchases.map(function(p) {
    const sup = suppliers[p.SupplierID];
    const lines = (itemsByPurchase[p.PurchaseID] || []).map(function(l) {
      const inv = invIndex[l.ItemID];
      return Object.assign({}, l, { ItemName: inv ? inv.ItemName : '—' });
    });
    const paid = (paymentsByPurchase[p.PurchaseID] || []).reduce((s, x) => s + Number(x.AmountPaid||0), 0);
    const total = Number(p.TotalAmount) || 0;
    return Object.assign({}, p, { SupplierName: sup ? sup.SupplierName : '—', Items: lines, AmountPaid: paid, Balance: total - paid });
  });
}
function getPurchasesPageData() {
  return { purchases: getPurchases(), suppliers: getSuppliers(), inventory: getInventory() };
}
function applyPurchaseLines_(purchaseId, items) {
  let total = 0;
  const invSheet = getSheet_('Inventory');
  const invIndex = indexBy_(sheetToObjects_('Inventory'), 'ItemID');
  items.forEach(function(line) {
    const qty = Number(line.Quantity) || 0;
    const cost = Number(line.UnitCost) || 0;
    const subtotal = qty * cost;
    total += subtotal;
    createRecord_('PurchaseItems', { PurchaseID: purchaseId, ItemID: line.ItemID, Quantity: qty, UnitCost: cost, Subtotal: subtotal });
    const inv = invIndex[line.ItemID];
    if (inv) {
      const newQty = (Number(inv.QuantityInStock) || 0) + qty;
      invSheet.getRange(inv._row, SCHEMA.Inventory.indexOf('QuantityInStock') + 1).setValue(newQty);
      invSheet.getRange(inv._row, SCHEMA.Inventory.indexOf('UnitCost') + 1).setValue(cost);
      inv.QuantityInStock = newQty;
    }
  });
  return total;
}
function reversePurchaseLines_(purchaseId) {
  const lines = sheetToObjects_('PurchaseItems').filter(l => String(l.PurchaseID) === String(purchaseId));
  const invSheet = getSheet_('Inventory');
  const invIndex = indexBy_(sheetToObjects_('Inventory'), 'ItemID');
  lines.forEach(function(l) {
    const inv = invIndex[l.ItemID];
    if (inv) {
      const newQty = Math.max(0, (Number(inv.QuantityInStock) || 0) - Number(l.Quantity));
      invSheet.getRange(inv._row, SCHEMA.Inventory.indexOf('QuantityInStock') + 1).setValue(newQty);
    }
  });
  lines.forEach(l => deleteRecord_('PurchaseItems', l.PurchaseItemID));
}
function addPurchase(header, items) {
  header.PurchaseDate = header.PurchaseDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  header.Status = header.Status || 'Received';
  const purchaseId = getNextId_('Purchases');
  header.PurchaseID = purchaseId;
  header.TotalAmount = applyPurchaseLines_(purchaseId, items);
  const r = createRecord_('Purchases', header);
  invalidateDashboardCache_();
  logAction_('create', 'Purchases', purchaseId, { total: header.TotalAmount });
  return r;
}
function updatePurchase(id, header, items) {
  reversePurchaseLines_(id);
  header.TotalAmount = applyPurchaseLines_(id, items);
  const r = updateRecord_('Purchases', id, header);
  invalidateDashboardCache_();
  logAction_('update', 'Purchases', id, header);
  return r;
}
function deletePurchase(id) {
  reversePurchaseLines_(id);
  const r = deleteRecord_('Purchases', id);
  invalidateDashboardCache_();
  logAction_('delete', 'Purchases', id, {});
  return r;
}

// ---------- SALES --------------------------------------------------------
function getSales() {
  const sales = sheetToObjects_('Sales');
  const customers = indexBy_(sheetToObjects_('Customers'), 'CustomerID');
  const invIndex = indexBy_(sheetToObjects_('Inventory'), 'ItemID');
  const itemsBySale = groupBy_(sheetToObjects_('SaleItems'), 'SaleID');
  const receiptsBySale = groupBy_(sheetToObjects_('Receipts'), 'SaleID');
  return sales.map(function(s) {
    const cust = customers[s.CustomerID];
    const lines = (itemsBySale[s.SaleID] || []).map(function(l) {
      const inv = invIndex[l.ItemID];
      const cost = inv ? Number(inv.UnitCost) || 0 : 0;
      const sub = Number(l.Subtotal) || 0;
      return Object.assign({}, l, { ItemName: inv ? inv.ItemName : '—', LineCost: cost * Number(l.Quantity || 0), LineProfit: sub - cost * Number(l.Quantity || 0) });
    });
    const received = (receiptsBySale[s.SaleID] || []).reduce((s2, x) => s2 + Number(x.AmountReceived||0), 0);
    const total = Number(s.TotalAmount) || 0;
    return Object.assign({}, s, { CustomerName: cust ? cust.CustomerName : '—', City: cust ? cust.City : '', State: cust ? cust.State : '', Items: lines, AmountReceived: received, Balance: total - received });
  });
}
function getSalesPageData() {
  return { sales: getSales(), customers: getCustomers(), inventory: getInventory() };
}
function applySaleLines_(saleId, items) {
  let total = 0;
  const invSheet = getSheet_('Inventory');
  const invIndex = indexBy_(sheetToObjects_('Inventory'), 'ItemID');
  items.forEach(function(line) {
    const qty = Number(line.Quantity) || 0;
    const inv = invIndex[line.ItemID];
    if (inv && Number(inv.QuantityInStock) < qty) throw new Error(`Insufficient stock for ${inv.ItemName} (have ${inv.QuantityInStock}, need ${qty})`);
    const price = Number(line.UnitPrice) || (inv ? Number(inv.UnitPrice) : 0);
    const subtotal = qty * price;
    total += subtotal;
    createRecord_('SaleItems', { SaleID: saleId, ItemID: line.ItemID, Quantity: qty, UnitPrice: price, Subtotal: subtotal });
    if (inv) {
      const newQty = (Number(inv.QuantityInStock) || 0) - qty;
      invSheet.getRange(inv._row, SCHEMA.Inventory.indexOf('QuantityInStock') + 1).setValue(newQty);
      inv.QuantityInStock = newQty;
    }
  });
  return total;
}
function reverseSaleLines_(saleId) {
  const lines = sheetToObjects_('SaleItems').filter(l => String(l.SaleID) === String(saleId));
  const invSheet = getSheet_('Inventory');
  const invIndex = indexBy_(sheetToObjects_('Inventory'), 'ItemID');
  lines.forEach(function(l) {
    const inv = invIndex[l.ItemID];
    if (inv) {
      const newQty = (Number(inv.QuantityInStock) || 0) + Number(l.Quantity);
      invSheet.getRange(inv._row, SCHEMA.Inventory.indexOf('QuantityInStock') + 1).setValue(newQty);
    }
  });
  lines.forEach(l => deleteRecord_('SaleItems', l.SaleItemID));
}
function addSale(header, items) {
  header.SaleDate = header.SaleDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  header.Status = header.Status || 'Completed';
  if (!header.Location) {
    const cust = indexBy_(sheetToObjects_('Customers'), 'CustomerID')[header.CustomerID];
    header.Location = cust ? cust.State : '';
  }
  const saleId = getNextId_('Sales');
  header.SaleID = saleId;
  header.TotalAmount = applySaleLines_(saleId, items);
  const r = createRecord_('Sales', header);
  invalidateDashboardCache_();
  logAction_('create', 'Sales', saleId, { total: header.TotalAmount });
  return r;
}
function updateSale(id, header, items) {
  reverseSaleLines_(id);
  header.TotalAmount = applySaleLines_(id, items);
  const r = updateRecord_('Sales', id, header);
  invalidateDashboardCache_();
  logAction_('update', 'Sales', id, header);
  return r;
}
function deleteSale(id) {
  reverseSaleLines_(id);
  const r = deleteRecord_('Sales', id);
  invalidateDashboardCache_();
  logAction_('delete', 'Sales', id, {});
  return r;
}

// ---------- RECEIPTS -----------------------------------------------------
function getReceipts() {
  const receipts = sheetToObjects_('Receipts');
  const sales = indexBy_(sheetToObjects_('Sales'), 'SaleID');
  const customers = indexBy_(sheetToObjects_('Customers'), 'CustomerID');
  return receipts.map(function(r) {
    const s = sales[r.SaleID];
    const cust = s ? customers[s.CustomerID] : null;
    return Object.assign({}, r, { SaleTotal: s ? s.TotalAmount : 0, CustomerName: cust ? cust.CustomerName : '—' });
  });
}
function getReceiptsPageData() { return { receipts: getReceipts(), sales: getSales() }; }
function addReceipt(data) {
  data.ReceiptDate = data.ReceiptDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  data.AmountReceived = Number(data.AmountReceived) || 0;
  const r = createRecord_('Receipts', data);
  invalidateDashboardCache_();
  logAction_('create', 'Receipts', r.ReceiptID, data);
  return r;
}
function updateReceipt(id, data) {
  if (data.AmountReceived !== undefined) data.AmountReceived = Number(data.AmountReceived);
  const r = updateRecord_('Receipts', id, data);
  invalidateDashboardCache_();
  logAction_('update', 'Receipts', id, data);
  return r;
}
function deleteReceipt(id) {
  const r = deleteRecord_('Receipts', id);
  invalidateDashboardCache_();
  logAction_('delete', 'Receipts', id, {});
  return r;
}

// ---------- PAYMENTS -----------------------------------------------------
function getPayments() {
  const payments = sheetToObjects_('Payments');
  const purchases = indexBy_(sheetToObjects_('Purchases'), 'PurchaseID');
  const suppliers = indexBy_(sheetToObjects_('Suppliers'), 'SupplierID');
  return payments.map(function(p) {
    const pur = purchases[p.PurchaseID];
    const sup = pur ? suppliers[pur.SupplierID] : null;
    return Object.assign({}, p, { PurchaseTotal: pur ? pur.TotalAmount : 0, SupplierName: sup ? sup.SupplierName : '—' });
  });
}
function getPaymentsPageData() { return { payments: getPayments(), purchases: getPurchases() }; }
function addPayment(data) {
  data.PaymentDate = data.PaymentDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  data.AmountPaid = Number(data.AmountPaid) || 0;
  const r = createRecord_('Payments', data);
  invalidateDashboardCache_();
  logAction_('create', 'Payments', r.PaymentID, data);
  return r;
}
function updatePayment(id, data) {
  if (data.AmountPaid !== undefined) data.AmountPaid = Number(data.AmountPaid);
  const r = updateRecord_('Payments', id, data);
  invalidateDashboardCache_();
  logAction_('update', 'Payments', id, data);
  return r;
}
function deletePayment(id) {
  const r = deleteRecord_('Payments', id);
  invalidateDashboardCache_();
  logAction_('delete', 'Payments', id, {});
  return r;
}

// ---------- USERS --------------------------------------------------------
function getUsers() {
  return sheetToObjects_('Users').map(u => {
    const clone = Object.assign({}, u);
    delete clone.Password;
    return clone;
  });
}
function addUser(data) {
  data.DateAdded = data.DateAdded || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  data.Status = data.Status || 'Active';
  const r = createRecord_('Users', data);
  logAction_('create', 'Users', r.UserID, { email: r.Email });
  return r;
}
function updateUser(id, data) {
  if (data.Password === '') delete data.Password;
  const r = updateRecord_('Users', id, data);
  logAction_('update', 'Users', id, data);
  return r;
}
function deleteUser(id) {
  const r = deleteRecord_('Users', id);
  logAction_('delete', 'Users', id, {});
  return r;
}

// ---------- SETTINGS -----------------------------------------------------
function getSettings() {
  const rows = sheetToObjects_('Settings');
  const map = {};
  rows.forEach(r => map[r.Key] = r.Value);
  const defaults = {
    CompanyName: 'Inventory Management System',
    Currency: 'USD',
    CurrencySymbol: '$',
    TaxRate: '0',
    LowStockThreshold: '10',
    InvoicePrefix: 'INV',
    PurchasePrefix: 'PO'
  };
  Object.keys(defaults).forEach(k => { if (!(k in map)) map[k] = defaults[k]; });
  return map;
}
function saveSettings(data) {
  const sheet = getSheet_('Settings');
  Object.keys(data).forEach(k => {
    const all = sheetToObjects_('Settings');
    const found = all.find(r => r.Key === k);
    if (found) sheet.getRange(found._row, 2).setValue(data[k]);
    else sheet.appendRow([k, data[k]]);
  });
  logAction_('settings', 'Settings', null, data);
  return getSettings();
}

// ---------- REPORTS ------------------------------------------------------
function getProfitLossReport(startDate, endDate) {
  const start = startDate ? new Date(startDate) : new Date('2000-01-01');
  const end = endDate ? new Date(endDate) : new Date('2100-01-01');
  const inRange = d => {
    const dd = d ? new Date(d) : null;
    return dd && dd >= start && dd <= end;
  };
  const sales = sheetToObjects_('Sales').filter(s => inRange(s.SaleDate));
  const saleItems = sheetToObjects_('SaleItems');
  const invIndex = indexBy_(sheetToObjects_('Inventory'), 'ItemID');
  const purchases = sheetToObjects_('Purchases').filter(p => inRange(p.PurchaseDate));
  const totalRevenue = sales.reduce((s, x) => s + (Number(x.TotalAmount)||0), 0);
  const totalCogs = sales.reduce((s, sale) => {
    const lines = saleItems.filter(l => String(l.SaleID) === String(sale.SaleID));
    return s + lines.reduce((ss, l) => {
      const inv = invIndex[l.ItemID];
      const cost = inv ? Number(inv.UnitCost)||0 : 0;
      return ss + cost * (Number(l.Quantity)||0);
    }, 0);
  }, 0);
  const totalPurchases = purchases.reduce((s, x) => s + (Number(x.TotalAmount)||0), 0);
  const grossProfit = totalRevenue - totalCogs;
  const netProfit = totalRevenue - totalPurchases;
  return {
    totalRevenue, totalCogs, totalPurchases, grossProfit, netProfit,
    margin: totalRevenue > 0 ? (grossProfit / totalRevenue) * 100 : 0,
    saleCount: sales.length, purchaseCount: purchases.length
  };
}
function getStockValuationReport() {
  const items = getInventory();
  let totalCost = 0, totalRetail = 0, totalPotentialProfit = 0;
  const byCategory = {};
  items.forEach(it => {
    totalCost += Number(it.StockValue) || 0;
    totalRetail += Number(it.PotentialRevenue) || 0;
    totalPotentialProfit += Number(it.PotentialProfit) || 0;
    const cat = it.Category || 'Uncategorized';
    if (!byCategory[cat]) byCategory[cat] = { qty: 0, cost: 0, retail: 0 };
    byCategory[cat].qty += Number(it.QuantityInStock) || 0;
    byCategory[cat].cost += Number(it.StockValue) || 0;
    byCategory[cat].retail += Number(it.PotentialRevenue) || 0;
  });
  return {
    totalCost, totalRetail, totalPotentialProfit, itemCount: items.length,
    byCategory: Object.keys(byCategory).map(k => ({ category: k, ...byCategory[k] }))
  };
}
function getLowStockReport() { return getInventory().filter(it => it.IsLowStock); }
function getSalesByPeriod(startDate, endDate) {
  const start = startDate ? new Date(startDate) : new Date('2000-01-01');
  const end = endDate ? new Date(endDate) : new Date('2100-01-01');
  const sales = sheetToObjects_('Sales').filter(s => {
    const d = s.SaleDate ? new Date(s.SaleDate) : null;
    return d && d >= start && d <= end;
  });
  const byDay = {};
  sales.forEach(s => {
    const key = s.SaleDate; // already "yyyy-MM-dd"
    byDay[key] = (byDay[key] || 0) + (Number(s.TotalAmount) || 0);
  });
  return Object.keys(byDay).sort().map(k => ({ date: k, value: byDay[k] }));
}

// ---------- AUDIT LOG READ ------------------------------------------------
function getAuditLog(limit) {
  limit = limit || 100;
  return sheetToObjects_('AuditLog')
    .sort((a, b) => new Date(b.Timestamp).getTime() - new Date(a.Timestamp).getTime())
    .slice(0, limit);
}

// ---------- OPTIONAL: SAMPLE DATA SEEDER ----------------------------------
function seedSampleData() {
  addUser({
    Name: 'Gogo Admin',
    Email: 'profilemail.info@gmail.com',
    Role: 'Admin',
    Status: 'Active',
    Password: 'Suresh@40'
  });

  const s1 = addSupplier({SupplierName:'BrightHive Innovations', ContactPerson:'Jon Reyes', Phone:'555-0101', Email:'jon@brighthive.com', City:'Austin', State:'Texas'});
  const s2 = addSupplier({SupplierName:'Apex Digital Systems', ContactPerson:'Mia Chen', Phone:'555-0102', Email:'mia@apexdigital.com', City:'San Jose', State:'California'});
  const s3 = addSupplier({SupplierName:'NexusTech Partners', ContactPerson:'Sam Patel', Phone:'555-0103', Email:'sam@nexustech.com', City:'Phoenix', State:'Arizona'});

  const i1 = addInventoryItem({ItemName:'Graphic Card RTX 4070', SKU:'GPU-001', Category:'Graphic Cards', SupplierID:s1.SupplierID, QuantityInStock:40, UnitCost:380, UnitPrice:520, ReorderLevel:10});
  const i2 = addInventoryItem({ItemName:'Motherboard Z790', SKU:'MB-002', Category:'Motherboards', SupplierID:s2.SupplierID, QuantityInStock:25, UnitCost:180, UnitPrice:260, ReorderLevel:8});
  const i3 = addInventoryItem({ItemName:'Ryzen 9 Processor', SKU:'CPU-003', Category:'Processors', SupplierID:s3.SupplierID, QuantityInStock:30, UnitCost:260, UnitPrice:370, ReorderLevel:10});
  const i4 = addInventoryItem({ItemName:'32GB Memory Kit', SKU:'MEM-004', Category:'Memory Chips', SupplierID:s1.SupplierID, QuantityInStock:60, UnitCost:70, UnitPrice:110, ReorderLevel:15});
  const i5 = addInventoryItem({ItemName:'NVMe SSD 1TB', SKU:'SSD-005', Category:'Storage', SupplierID:s2.SupplierID, QuantityInStock:35, UnitCost:60, UnitPrice:95, ReorderLevel:12});
  const i6 = addInventoryItem({ItemName:'Power Supply 850W', SKU:'PSU-006', Category:'Power', SupplierID:s3.SupplierID, QuantityInStock:18, UnitCost:90, UnitPrice:140, ReorderLevel:6});

  const c1 = addCustomer({CustomerName:'Summit IT Services', Phone:'555-0201', Email:'info@summitit.com', City:'Miami', State:'Florida'});
  const c2 = addCustomer({CustomerName:'Sunbelt Digital Solutions', Phone:'555-0202', Email:'info@sunbelt.com', City:'Houston', State:'Texas'});
  const c3 = addCustomer({CustomerName:'Evergreen Tech Solutions', Phone:'555-0203', Email:'info@evergreen.com', City:'Los Angeles', State:'California'});

  addPurchase({SupplierID:s1.SupplierID}, [{ItemID:i1.ItemID, Quantity:20, UnitCost:380}]);
  addPurchase({SupplierID:s2.SupplierID}, [{ItemID:i2.ItemID, Quantity:15, UnitCost:180}]);
  addSale({CustomerID:c1.CustomerID, Location:'Florida'}, [{ItemID:i1.ItemID, Quantity:5, UnitPrice:520}]);
  addSale({CustomerID:c2.CustomerID, Location:'Texas'}, [{ItemID:i3.ItemID, Quantity:3, UnitPrice:370}]);
  addSale({CustomerID:c3.CustomerID, Location:'California'}, [{ItemID:i4.ItemID, Quantity:10, UnitPrice:110}]);

  return 'Sample data added. Login: profilemail.info@gmail.com / Suresh@40';
}

// ---------- DASHBOARD AGGREGATION (cached 30s) ----------------------------
function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dashboard_v1');
  if (cached) return JSON.parse(cached);

  const sales         = sheetToObjects_('Sales');
  const purchases     = sheetToObjects_('Purchases');
  const receipts      = sheetToObjects_('Receipts');
  const payments      = sheetToObjects_('Payments');
  const saleItems     = sheetToObjects_('SaleItems');
  const purchaseItems = sheetToObjects_('PurchaseItems');
  const invRows       = sheetToObjects_('Inventory');
  const inventory     = indexBy_(invRows, 'ItemID');
  const customers     = indexBy_(sheetToObjects_('Customers'), 'CustomerID');
  const suppliers     = indexBy_(sheetToObjects_('Suppliers'), 'SupplierID');
  const purchasesById = indexBy_(purchases, 'PurchaseID');

  const sum = (arr, field) => arr.reduce((s, r) => s + (Number(r[field])||0), 0);

  const totalSales      = sum(sales, 'TotalAmount');
  const totalPurchases  = sum(purchases, 'TotalAmount');
  const totalReceived   = sum(receipts, 'AmountReceived');
  const totalPaid       = sum(payments, 'AmountPaid');
  const netProfit       = totalSales - totalPurchases;
  const totalReceivable = totalSales - totalReceived;
  const totalPayable    = totalPurchases - totalPaid;

  const totalCogs = saleItems.reduce((s, l) => {
    const inv = inventory[l.ItemID];
    const cost = inv ? Number(inv.UnitCost)||0 : 0;
    return s + cost * (Number(l.Quantity)||0);
  }, 0);
  const grossProfit = totalSales - totalCogs;

  // sales trend
  const trendMap = {};
  sales.forEach(s => {
    const d = s.SaleDate || 'Unknown';
    trendMap[d] = (trendMap[d]||0) + (Number(s.TotalAmount)||0);
  });
  const salesTrend = Object.keys(trendMap).sort().map(d => ({ date: d, value: trendMap[d] }));

  // top customers
  const custTotals = {};
  sales.forEach(s => custTotals[s.CustomerID] = (custTotals[s.CustomerID]||0) + (Number(s.TotalAmount)||0));
  const topCustomers = Object.keys(custTotals)
    .map(cid => ({ name: customers[cid] ? customers[cid].CustomerName : cid, value: custTotals[cid] }))
    .sort((a,b) => b.value - a.value).slice(0,10);

  // purchase by location
  const purByState = {};
  purchases.forEach(p => {
    const sup = suppliers[p.SupplierID];
    const state = sup ? (sup.State||'Unknown') : 'Unknown';
    purByState[state] = (purByState[state]||0) + (Number(p.TotalAmount)||0);
  });
  const purchaseByLocation = Object.keys(purByState).map(k => ({ label: k, value: purByState[k] }));

  // purchase by category × year
  const purByCategoryYear = {};
  purchaseItems.forEach(l => {
    const purchase = purchasesById[l.PurchaseID];
    const year = purchase && purchase.PurchaseDate ? purchase.PurchaseDate.substring(0,4) : 'Unknown';
    const inv = inventory[l.ItemID];
    const cat = inv ? (inv.Category||'Uncategorized') : 'Uncategorized';
    if (!purByCategoryYear[year]) purByCategoryYear[year] = {};
    purByCategoryYear[year][cat] = (purByCategoryYear[year][cat]||0) + (Number(l.Subtotal)||0);
  });
  const yearSet = {}, catSet = {};
  Object.keys(purByCategoryYear).forEach(y => {
    yearSet[y] = true;
    Object.keys(purByCategoryYear[y]).forEach(c => catSet[c] = true);
  });
  const years = Object.keys(yearSet).sort();
  const categories = Object.keys(catSet);
  const purMatrix = years.map(y => categories.map(c => purByCategoryYear[y][c] || 0));

  // sales by location
  const salesByLocationMap = {};
  sales.forEach(s => {
    const loc = s.Location || 'Unknown';
    salesByLocationMap[loc] = (salesByLocationMap[loc]||0) + (Number(s.TotalAmount)||0);
  });
  const salesByLocation = Object.keys(salesByLocationMap).map(k => ({ label: k, value: salesByLocationMap[k] }));

  // sales by category
  const salesByCategoryMap = {};
  saleItems.forEach(l => {
    const inv = inventory[l.ItemID];
    const cat = inv ? (inv.Category||'Uncategorized') : 'Uncategorized';
    salesByCategoryMap[cat] = (salesByCategoryMap[cat]||0) + (Number(l.Subtotal)||0);
  });
  const salesByCategory = Object.keys(salesByCategoryMap).map(k => ({ label: k, value: salesByCategoryMap[k] }));

  // sales by city
  const cityMap = {};
  sales.forEach(s => {
    const cust = customers[s.CustomerID];
    const city = cust ? (cust.City||'Unknown') : 'Unknown';
    cityMap[city] = (cityMap[city]||0) + (Number(s.TotalAmount)||0);
  });
  const salesByCity = Object.keys(cityMap).map(k => ({ label: k, value: cityMap[k] })).sort((a,b) => b.value - a.value);

  const topSalesLocation = salesByLocation.sort((a,b) => b.value - a.value)[0];
  const topSalesLocationName = topSalesLocation ? topSalesLocation.label : '—';

  const qtyByItem = {};
  saleItems.forEach(l => qtyByItem[l.ItemID] = (qtyByItem[l.ItemID]||0) + (Number(l.Quantity)||0));
  const topItemId = Object.keys(qtyByItem).sort((a,b) => qtyByItem[b] - qtyByItem[a])[0];
  const topSellingItem = topItemId && inventory[topItemId] ? inventory[topItemId].ItemName : '—';

  const lowStock = invRows.filter(i => (Number(i.QuantityInStock)||0) <= (Number(i.ReorderLevel)||0)).length;
  const inventoryValue = invRows.reduce((s, i) => s + (Number(i.QuantityInStock)||0) * (Number(i.UnitCost)||0), 0);

  const result = {
    totalSales, totalPurchases, netProfit, grossProfit, totalCogs,
    totalReceivable, totalPayable,
    topSalesLocation: topSalesLocationName,
    topSellingItem,
    salesTrend, topCustomers,
    purchaseByLocation,
    purchaseByCategoryYear: { years, categories, matrix: purMatrix },
    salesByLocation, salesByCategory, salesByCity,
    lowStock, inventoryCount: invRows.length, inventoryValue
  };

  cache.put('dashboard_v1', JSON.stringify(result), 30);
  return result;
}
