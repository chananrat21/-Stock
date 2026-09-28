// Google Apps Script template for Google Sheets integration
// 1) Open Google Sheets
// 2) Extensions > Apps Script
// 3) Paste this code
// 4) Deploy > New deployment > Web app
// 5) Set access to "Anyone" or "Anyone with Google account"

const SHEET_NAMES = {
  PRODUCTS: 'products',
  MOVEMENTS: 'stock_movements',
  SUPPLIERS: 'suppliers',
  SETTINGS: 'settings'
};

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function ensureHeaders() {
  const products = getSheet(SHEET_NAMES.PRODUCTS);
  const movements = getSheet(SHEET_NAMES.MOVEMENTS);
  const suppliers = getSheet(SHEET_NAMES.SUPPLIERS);
  const settings = getSheet(SHEET_NAMES.SETTINGS);

  const productHeaders = ['product_id', 'name', 'category', 'unit', 'qty', 'min_stock', 'price_per_unit', 'supplier', 'image_url', 'mfg_date', 'expiry_date', 'status', 'updated_at'];
  const movementHeaders = ['movement_id', 'product_id', 'movement_type', 'qty', 'reason', 'note', 'created_by', 'created_at'];
  const supplierHeaders = ['supplier_id', 'name', 'phone', 'email', 'updated_at'];
  const settingsHeaders = ['key', 'value', 'updated_at'];

  setHeadersIfNeeded(products, productHeaders);
  setHeadersIfNeeded(movements, movementHeaders);
  setHeadersIfNeeded(suppliers, supplierHeaders);
  setHeadersIfNeeded(settings, settingsHeaders);
}

function setHeadersIfNeeded(sheet, headers) {
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = existing && existing.some(v => v && String(v).trim());

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function getRowsAsObjects(sheetName) {
  ensureHeaders();
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  return values.slice(1).filter(row => row.some(cell => cell !== '')).map(row => {
    const obj = {};
    headers.forEach((key, index) => {
      obj[key] = row[index] ?? '';
    });
    return obj;
  });
}

function generateId(prefix) {
  const stamp = new Date().getTime();
  return `${prefix}_${stamp}`;
}

function normalizeProductPayload(payload) {
  const now = new Date().toISOString();
  return {
    product_id: payload.product_id || generateId('P'),
    name: payload.name || '',
    category: payload.category || '',
    unit: payload.unit || '',
    qty: Number(payload.qty || 0),
    min_stock: Number(payload.min_stock || 0),
    price_per_unit: Number(payload.price_per_unit || 0),
    supplier: payload.supplier || '',
    image_url: payload.image_url || '',
    mfg_date: payload.mfg_date || '',
    expiry_date: payload.expiry_date || '',
    status: payload.status || 'OK',
    updated_at: now,
  };
}

function responseJson(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const type = (e && e.parameter && e.parameter.type) || 'products';

  if (type === 'products') {
    return responseJson(getRowsAsObjects(SHEET_NAMES.PRODUCTS));
  }

  if (type === 'stock_movements') {
    return responseJson(getRowsAsObjects(SHEET_NAMES.MOVEMENTS));
  }

  if (type === 'dashboard') {
    const products = getRowsAsObjects(SHEET_NAMES.PRODUCTS);
    const totalQty = products.reduce((sum, p) => sum + Number(p.qty || 0), 0);
    const lowStock = products.filter(p => Number(p.qty || 0) <= Number(p.min_stock || 0)).length;
    const zeroStock = products.filter(p => Number(p.qty || 0) <= 0).length;

    return responseJson({
      totalProducts: products.length,
      totalQty,
      lowStock,
      zeroStock,
      products
    });
  }

  if (type === 'settings') {
    return responseJson(getRowsAsObjects(SHEET_NAMES.SETTINGS));
  }

  return responseJson({ success: false, message: 'Unknown type' });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const type = payload.type || 'addProduct';

    if (type === 'addProduct') {
      ensureHeaders();
      const row = normalizeProductPayload(payload);
      const sheet = getSheet(SHEET_NAMES.PRODUCTS);
      const existing = sheet.getDataRange().getValues();
      const hasProductId = existing.slice(1).some(r => String(r[0]) === String(row.product_id));

      if (hasProductId) {
        const targetRow = existing.findIndex(r => String(r[0]) === String(row.product_id));
        if (targetRow > 0) {
          const values = [
            row.product_id,
            row.name,
            row.category,
            row.unit,
            row.qty,
            row.min_stock,
            row.price_per_unit,
            row.supplier,
            row.image_url,
            row.mfg_date,
            row.expiry_date,
            row.status,
            row.updated_at
          ];
          sheet.getRange(targetRow + 1, 1, 1, values.length).setValues([values]);
        }
      } else {
        sheet.appendRow([
          row.product_id,
          row.name,
          row.category,
          row.unit,
          row.qty,
          row.min_stock,
          row.price_per_unit,
          row.supplier,
          row.image_url,
          row.mfg_date,
          row.expiry_date,
          row.status,
          row.updated_at
        ]);
      }

      return responseJson({ success: true, product: row });
    }

    if (type === 'addMovement') {
      ensureHeaders();
      const movementSheet = getSheet(SHEET_NAMES.MOVEMENTS);
      const movementId = generateId('M');
      const now = new Date().toISOString();
      const row = [
        movementId,
        payload.product_id || '',
        payload.movement_type || 'IN',
        Number(payload.qty || 0),
        payload.reason || '',
        payload.note || '',
        payload.created_by || 'system',
        now
      ];
      movementSheet.appendRow(row);

      if (payload.product_id) {
        const products = getSheet(SHEET_NAMES.PRODUCTS);
        const productValues = products.getDataRange().getValues();
        const headers = productValues[0];
        const index = headers.indexOf('product_id');

        for (let i = 1; i < productValues.length; i++) {
          if (String(productValues[i][index]) === String(payload.product_id)) {
            const qtyIndex = headers.indexOf('qty');
            const minIndex = headers.indexOf('min_stock');
            const currentQty = Number(productValues[i][qtyIndex] || 0);
            const delta = Number(payload.qty || 0);
            const nextQty = payload.movement_type === 'OUT' ? currentQty - delta : currentQty + delta;

            products.getRange(i + 1, qtyIndex + 1).setValue(nextQty);
            products.getRange(i + 1, headers.indexOf('updated_at') + 1).setValue(now);
            products.getRange(i + 1, headers.indexOf('status') + 1).setValue(nextQty <= 0 ? 'OUT' : nextQty <= Number(productValues[i][minIndex] || 0) ? 'LOW' : 'OK');
            break;
          }
        }
      }

      return responseJson({ success: true, movementId, created_at: now });
    }

    if (type === 'updateProduct') {
      ensureHeaders();
      const row = normalizeProductPayload(payload);
      const sheet = getSheet(SHEET_NAMES.PRODUCTS);
      const values = sheet.getDataRange().getValues();
      const headers = values[0];
      const idIndex = headers.indexOf('product_id');

      for (let i = 1; i < values.length; i++) {
        if (String(values[i][idIndex]) === String(row.product_id)) {
          const target = [
            row.product_id,
            row.name,
            row.category,
            row.unit,
            row.qty,
            row.min_stock,
            row.price_per_unit,
            row.supplier,
            row.image_url,
            row.mfg_date,
            row.expiry_date,
            row.status,
            row.updated_at
          ];
          sheet.getRange(i + 1, 1, 1, target.length).setValues([target]);
          return responseJson({ success: true, product: row });
        }
      }

      return responseJson({ success: false, message: 'Product not found' });
    }

    return responseJson({ success: false, message: 'Unsupported action' });
  } catch (err) {
    return responseJson({ success: false, message: err && err.message ? err.message : 'Server error' });
  }
}

function testScript() {
  const payload = {
    type: 'addProduct',
    name: 'Sample Product',
    category: 'Fertilizer',
    unit: 'Bag',
    qty: 20,
    min_stock: 5,
    price_per_unit: 500,
    supplier: 'Supplier A'
  };

  console.log('Sample payload created:', payload);
}
