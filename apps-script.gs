/**
 * CREATIVE VISION — Shop backend (Google Apps Script)
 * -----------------------------------------------------
 * This script turns a Google Sheet into a simple free backend for
 * shop.html (products + checkout), quote.html (quotation requests),
 * enquire.html (general enquiries), and admin.html (manage products,
 * view orders/quotations/enquiries).
 *
 * SETUP: see SETUP-GUIDE.md for full step-by-step instructions.
 * You only need to edit the one line below.
 *
 * IMPORTANT — if you're updating from an earlier version of this script:
 * your Orders tab needs two new column headers added by hand, if you
 * don't already have them.
 * In cell I1 of the Orders tab, type: Type
 * In cell J1 of the Orders tab, type: Notes
 * (Columns A–H stay exactly as they were: Timestamp, OrderID, CustomerName,
 * Phone, Address, Items, Total, Status — Type is column 9, Notes is column 10.)
 */

const ADMIN_KEY = "Admin@01199010"; // change this, then copy the
                                                   // SAME value into admin.html and shop.html

// ---------------------------------------------------------------------
// ONE-TIME SETUP HELPER — run this once manually, see SETUP-GUIDE.md
// ---------------------------------------------------------------------
// Select "authorizeDrive" in the function dropdown above (next to the Run
// button), click Run, and approve the Google Drive permission when asked.
// This is required once so the web app is allowed to save uploaded photos
// to your Drive. You only need to do this a single time.
function authorizeDrive() {
  const folder = DriveApp.createFolder('Creative Vision Authorization Test — safe to delete');
  folder.setTrashed(true); // cleans up after itself
}

// ---------------------------------------------------------------------
// Web app entry points
// ---------------------------------------------------------------------

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'products') {
    return jsonResponse(getProducts());
  }

  if (action === 'orders') {
    if (e.parameter.key !== ADMIN_KEY) {
      return jsonResponse({ error: 'Unauthorized' });
    }
    return jsonResponse(getOrders());
  }

  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'Invalid request body' });
  }

  if (data.action === 'addProduct') {
    if (data.key !== ADMIN_KEY) return jsonResponse({ error: 'Unauthorized' });
    const result = addProduct(data);
    return jsonResponse(result);
  }

  if (data.action === 'updateProduct') {
    if (data.key !== ADMIN_KEY) return jsonResponse({ error: 'Unauthorized' });
    const result = updateProduct(data);
    return jsonResponse(result);
  }

  if (data.action === 'deleteProduct') {
    if (data.key !== ADMIN_KEY) return jsonResponse({ error: 'Unauthorized' });
    deleteProduct(data.id);
    return jsonResponse({ success: true });
  }

  if (data.action === 'updateStock') {
    if (data.key !== ADMIN_KEY) return jsonResponse({ error: 'Unauthorized' });
    updateStock(data.id, data.inStock);
    return jsonResponse({ success: true });
  }

  if (data.action === 'placeOrder') {
    const orderId = addOrder(data, 'Order');
    return jsonResponse({ success: true, orderId: orderId });
  }

  if (data.action === 'placeQuotation') {
    const orderId = addOrder(data, 'Quotation');
    return jsonResponse({ success: true, orderId: orderId });
  }

  if (data.action === 'placeEnquiry') {
    const orderId = addOrder(data, 'Enquiry');
    return jsonResponse({ success: true, orderId: orderId });
  }

  if (data.action === 'updateOrderStatus') {
    if (data.key !== ADMIN_KEY) return jsonResponse({ error: 'Unauthorized' });
    updateOrderStatus(data.orderId, data.status);
    return jsonResponse({ success: true });
  }

  return jsonResponse({ error: 'Unknown action' });
}

// ---------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" not found. Check tab names match exactly: Products, Orders`);
  return sheet;
}

function sheetToObjects(sheet) {
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  return rows
    .filter(r => r.join('') !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
}

function findRowIndexById(sheet, id) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) return i + 1; // 1-based sheet row number
  }
  return -1;
}

// ---------------------------------------------------------------------
// Products — sheet tab "Products", columns:
// ID | Name | Category | Price | ImageURL | Description | InStock
// ---------------------------------------------------------------------

function getProducts() {
  return sheetToObjects(getSheet('Products'));
}

function addProduct(data) {
  const sheet = getSheet('Products');
  const id = 'P' + new Date().getTime();
  let imageUrl = data.imageUrl || '';
  let imageWarning = null;

  if (data.imageData) {
    try {
      imageUrl = saveImageToDrive(data.imageData, data.imageName);
    } catch (err) {
      imageWarning = 'Image upload failed: ' + err.message;
      imageUrl = data.imageUrl || '';
    }
  }

  sheet.appendRow([
    id,
    data.name,
    data.category,
    Number(data.price) || 0,
    imageUrl,
    data.description || '',
    'Yes'
  ]);

  const result = { success: true, id: id };
  if (imageWarning) result.imageWarning = imageWarning;
  return result;
}

function updateProduct(data) {
  const sheet = getSheet('Products');
  const rowNum = findRowIndexById(sheet, data.id);
  if (rowNum === -1) return { error: 'Product not found' };

  let imageWarning = null;
  // Column order: ID(1) Name(2) Category(3) Price(4) ImageURL(5) Description(6) InStock(7)
  if (data.name) sheet.getRange(rowNum, 2).setValue(data.name);
  if (data.category) sheet.getRange(rowNum, 3).setValue(data.category);
  if (data.price) sheet.getRange(rowNum, 4).setValue(Number(data.price) || 0);

  if (data.imageData) {
    try {
      const imageUrl = saveImageToDrive(data.imageData, data.imageName);
      sheet.getRange(rowNum, 5).setValue(imageUrl);
    } catch (err) {
      imageWarning = 'Image upload failed: ' + err.message;
    }
  } else if (typeof data.imageUrl === 'string' && data.imageUrl.trim() !== '') {
    sheet.getRange(rowNum, 5).setValue(data.imageUrl.trim());
  }

  if (typeof data.description === 'string') sheet.getRange(rowNum, 6).setValue(data.description);

  const result = { success: true };
  if (imageWarning) result.imageWarning = imageWarning;
  return result;
}

// Saves a base64-encoded image (from the admin upload form) into a Drive
// folder, makes it publicly viewable, and returns a URL that works as an
// <img src> on the shop page.
function saveImageToDrive(base64Data, fileName) {
  const folderName = 'Creative Vision Product Images';
  let folder;
  const existingFolders = DriveApp.getFoldersByName(folderName);
  folder = existingFolders.hasNext() ? existingFolders.next() : DriveApp.createFolder(folderName);

  const commaIndex = base64Data.indexOf(',');
  const cleanBase64 = commaIndex >= 0 ? base64Data.substring(commaIndex + 1) : base64Data;
  const bytes = Utilities.base64Decode(cleanBase64);
  const blob = Utilities.newBlob(bytes, 'image/jpeg', fileName || ('product-' + new Date().getTime() + '.jpg'));

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1000`;
}

function deleteProduct(id) {
  const sheet = getSheet('Products');
  const rowNum = findRowIndexById(sheet, id);
  if (rowNum !== -1) sheet.deleteRow(rowNum);
}

function updateStock(id, inStock) {
  const sheet = getSheet('Products');
  const rowNum = findRowIndexById(sheet, id);
  if (rowNum !== -1) sheet.getRange(rowNum, 7).setValue(inStock); // column 7 = InStock
}

// ---------------------------------------------------------------------
// Orders — sheet tab "Orders", columns:
// Timestamp | OrderID | CustomerName | Phone | Address | Items | Total | Status | Type | Notes
// Type is "Order" (shop.html checkout), "Quotation" (quote.html), or "Enquiry" (enquire.html)
// ---------------------------------------------------------------------

function getOrders() {
  return sheetToObjects(getSheet('Orders'));
}

function addOrder(data, type) {
  const sheet = getSheet('Orders');
  let prefix = 'ORD';
  if (type === 'Quotation') prefix = 'QUO';
  if (type === 'Enquiry') prefix = 'ENQ';
  const orderId = prefix + new Date().getTime();
  sheet.appendRow([
    new Date(),
    orderId,
    data.customerName || '',
    data.phone || '',
    data.address || '',
    JSON.stringify(data.items || []),
    Number(data.total) || 0,
    'New',
    type,
    data.notes || ''
  ]);
  return orderId;
}

function updateOrderStatus(orderId, status) {
  const sheet = getSheet('Orders');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === orderId) {
      sheet.getRange(i + 1, 8).setValue(status); // column 8 = Status
      break;
    }
  }
}
