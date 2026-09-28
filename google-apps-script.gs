/**
 * Google Apps Script backend for Stock Management System
 * 
 * วิธีติดตั้ง:
 * 1) เปิด Google Sheets สร้างชีตใหม่
 * 2) ไปที่เมนู ส่วนขยาย (Extensions) > Apps Script
 * 3) ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดไฟล์นี้ลงไป
 * 4) กดปุ่ม "ทำให้ใช้งานได้" (Deploy) > "การทำให้ใช้งานได้รายการใหม่" (New deployment)
 * 5) เลือกประเภทเป็น "เว็บแอป" (Web app)
 * 6) กำหนดค่า:
 *    - คำอธิบาย: Stock API
 *    - ดำเนินการในฐานะ: ฉัน (Me)
 *    - ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน (Anyone)  <-- สำคัญมาก! ต้องเลือก Anyone
 * 7) กด "ทำให้ใช้งานได้" (Deploy) และคัดลอก URL ของเว็บแอป (ลงท้ายด้วย /exec)
 * 8) นำ URL ไปใส่ในเมนู "⚙ ตั้งค่า" บนหน้าเว็บระบบสต๊อก
 */

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

  const productHeaders = [
    'product_id', 'name', 'category', 'unit', 'qty', 
    'min_stock', 'price_per_unit', 'supplier', 'image_url', 
    'mfg_date', 'expiry_date', 'status', 'updated_at'
  ];
  const movementHeaders = [
    'movement_id', 'product_id', 'movement_type', 'qty', 
    'reason', 'note', 'created_by', 'created_at'
  ];
  const supplierHeaders = ['supplier_id', 'name', 'phone', 'email', 'updated_at'];
  const settingsHeaders = ['key', 'value', 'updated_at'];

  setHeadersIfNeeded(products, productHeaders);
  setHeadersIfNeeded(movements, movementHeaders);
  setHeadersIfNeeded(suppliers, supplierHeaders);
  setHeadersIfNeeded(settings, settingsHeaders);
}

function setHeadersIfNeeded(sheet, headers) {
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const hasHeaders = existing && existing.some(v => v && String(v).trim());

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e9');
    sheet.setFrozenRows(1);
  }
}

function getRowsAsObjects(sheetName) {
  ensureHeaders();
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((key, index) => {
        let val = row[index] ?? '';
        if (val instanceof Date) {
          val = val.toISOString().slice(0, 10);
        }
        obj[key] = val;
      });
      return obj;
    });
}

function generateId(prefix) {
  return `${prefix}_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
}

function normalizeProductPayload(payload) {
  const now = new Date().toISOString();
  const qty = Number(payload.qty ?? payload.current ?? 0);
  const minStock = Number(payload.min_stock ?? 5);
  let status = payload.status;
  if (!status) {
    status = qty <= 0 ? 'OUT' : (qty <= minStock ? 'LOW' : 'OK');
  }

  return {
    product_id: String(payload.product_id || payload.id || generateId('P')),
    name: String(payload.name || '').trim(),
    category: String(payload.category || 'ทั่วไป').trim(),
    unit: String(payload.unit || 'ชิ้น').trim(),
    qty: Math.max(0, qty),
    min_stock: Math.max(0, minStock),
    price_per_unit: Number(payload.price_per_unit ?? payload.price ?? 0),
    supplier: String(payload.supplier || '').trim(),
    image_url: String(payload.image_url || payload.img || '').trim(),
    mfg_date: String(payload.mfg_date || payload.mfg || '').trim(),
    expiry_date: String(payload.expiry_date || payload.expiry || '').trim(),
    status: status,
    updated_at: now,
  };
}

function responseJson(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    ensureHeaders();
    const type = (e && e.parameter && e.parameter.type) || 'products';

    if (type === 'ping' || type === 'test') {
      return responseJson({
        success: true,
        message: 'ระบบเชื่อมต่อ Google Apps Script สำเร็จ',
        timestamp: new Date().toISOString(),
        version: '2.0.0'
      });
    }

    if (type === 'products') {
      return responseJson(getRowsAsObjects(SHEET_NAMES.PRODUCTS));
    }

    if (type === 'stock_movements' || type === 'movements') {
      return responseJson(getRowsAsObjects(SHEET_NAMES.MOVEMENTS));
    }

    if (type === 'suppliers') {
      return responseJson(getRowsAsObjects(SHEET_NAMES.SUPPLIERS));
    }

    if (type === 'settings') {
      return responseJson(getRowsAsObjects(SHEET_NAMES.SETTINGS));
    }

    if (type === 'dashboard') {
      const products = getRowsAsObjects(SHEET_NAMES.PRODUCTS);
      const movements = getRowsAsObjects(SHEET_NAMES.MOVEMENTS);
      const totalQty = products.reduce((sum, p) => sum + Number(p.qty || 0), 0);
      const totalValue = products.reduce((sum, p) => sum + (Number(p.qty || 0) * Number(p.price_per_unit || 0)), 0);
      const lowStock = products.filter(p => Number(p.qty || 0) > 0 && Number(p.qty || 0) <= Number(p.min_stock || 5)).length;
      const zeroStock = products.filter(p => Number(p.qty || 0) <= 0).length;

      return responseJson({
        success: true,
        totalProducts: products.length,
        totalQty,
        totalValue,
        lowStock,
        zeroStock,
        recentMovements: movements.slice(-10).reverse(),
        products
      });
    }

    return responseJson({ success: false, message: `ไม่พบประเภทคำขอ: ${type}` });
  } catch (err) {
    return responseJson({ success: false, message: err.message || 'เกิดข้อผิดพลาดในการดึงข้อมูล' });
  }
}

function doPost(e) {
  try {
    ensureHeaders();
    let payload = {};

    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const type = payload.type || 'addProduct';

    // 1) เพิ่มหรืออัปเดตสินค้า (Add or Upsert Product)
    if (type === 'addProduct') {
      const row = normalizeProductPayload(payload);
      if (!row.name) {
        return responseJson({ success: false, message: 'กรุณาระบุชื่อสินค้า' });
      }

      const sheet = getSheet(SHEET_NAMES.PRODUCTS);
      const existing = sheet.getDataRange().getValues();
      const targetRow = existing.slice(1).findIndex(r => String(r[0]) === String(row.product_id));

      const rowValues = [
        row.product_id, row.name, row.category, row.unit, row.qty,
        row.min_stock, row.price_per_unit, row.supplier, row.image_url,
        row.mfg_date, row.expiry_date, row.status, row.updated_at
      ];

      if (targetRow >= 0) {
        sheet.getRange(targetRow + 2, 1, 1, rowValues.length).setValues([rowValues]);
      } else {
        sheet.appendRow(rowValues);
      }

      // บันทึกความเคลื่อนไหวเริ่มต้นถ้ามียอดตั้งต้น
      if (payload.initial_movement && row.qty > 0 && targetRow < 0) {
        const movementSheet = getSheet(SHEET_NAMES.MOVEMENTS);
        movementSheet.appendRow([
          generateId('M'),
          row.product_id,
          'IN',
          row.qty,
          'ยอดเริ่มต้น',
          'เพิ่มสินค้าใหม่เข้าคลัง',
          payload.created_by || 'admin',
          row.updated_at
        ]);
      }

      return responseJson({ success: true, product: row, message: 'บันทึกสินค้าเรียบร้อย' });
    }

    // 2) อัปเดตข้อมูลสินค้า (Update Product details)
    if (type === 'updateProduct') {
      const pid = String(payload.product_id || payload.id || '');
      if (!pid) return responseJson({ success: false, message: 'ไม่พบรหัสสินค้า (product_id)' });

      const sheet = getSheet(SHEET_NAMES.PRODUCTS);
      const values = sheet.getDataRange().getValues();
      const headers = values[0];
      const idIdx = headers.indexOf('product_id');

      for (let i = 1; i < values.length; i++) {
        if (String(values[i][idIdx]) === pid) {
          const now = new Date().toISOString();
          const currentQty = Number(payload.qty !== undefined ? payload.qty : values[i][headers.indexOf('qty')] || 0);
          const minStock = Number(payload.min_stock !== undefined ? payload.min_stock : values[i][headers.indexOf('min_stock')] || 5);
          const status = currentQty <= 0 ? 'OUT' : (currentQty <= minStock ? 'LOW' : 'OK');

          const updated = [
            pid,
            payload.name !== undefined ? payload.name : values[i][headers.indexOf('name')],
            payload.category !== undefined ? payload.category : values[i][headers.indexOf('category')],
            payload.unit !== undefined ? payload.unit : values[i][headers.indexOf('unit')],
            currentQty,
            minStock,
            Number(payload.price_per_unit !== undefined ? payload.price_per_unit : values[i][headers.indexOf('price_per_unit')] || 0),
            payload.supplier !== undefined ? payload.supplier : values[i][headers.indexOf('supplier')],
            payload.image_url !== undefined ? payload.image_url : values[i][headers.indexOf('image_url')],
            payload.mfg_date !== undefined ? payload.mfg_date : values[i][headers.indexOf('mfg_date')],
            payload.expiry_date !== undefined ? payload.expiry_date : values[i][headers.indexOf('expiry_date')],
            status,
            now
          ];

          sheet.getRange(i + 1, 1, 1, updated.length).setValues([updated]);
          return responseJson({ success: true, message: 'แก้ไขข้อมูลสินค้าสำเร็จ', product_id: pid });
        }
      }

      return responseJson({ success: false, message: 'ไม่พบรายการสินค้าที่ต้องการแก้ไข' });
    }

    // 3) ลบสินค้า (Delete Product)
    if (type === 'deleteProduct' || type === 'removeProduct') {
      const pid = String(payload.product_id || payload.id || '');
      if (!pid) return responseJson({ success: false, message: 'ไม่พบรหัสสินค้าที่ต้องการลบ' });

      const sheet = getSheet(SHEET_NAMES.PRODUCTS);
      const values = sheet.getDataRange().getValues();
      const idIdx = values[0].indexOf('product_id');

      for (let i = 1; i < values.length; i++) {
        if (String(values[i][idIdx]) === pid) {
          sheet.deleteRow(i + 1);
          return responseJson({ success: true, message: 'ลบรายการสินค้าเรียบร้อยแล้ว', product_id: pid });
        }
      }

      return responseJson({ success: false, message: 'ไม่พบรายการสินค้าที่ต้องการลบในระบบ' });
    }

    // 4) บันทึกความเคลื่อนไหว (Add Movement: IN, OUT, ADJUST)
    if (type === 'addMovement') {
      const movementSheet = getSheet(SHEET_NAMES.MOVEMENTS);
      const movementId = generateId('M');
      const now = new Date().toISOString();
      const moveType = String(payload.movement_type || payload.type || 'IN').toUpperCase();
      const qtyDelta = Math.abs(Number(payload.qty || 0));
      const pid = String(payload.product_id || payload.pid || '');

      if (!pid) return responseJson({ success: false, message: 'กรุณาระบุรหัสสินค้า' });
      if (qtyDelta <= 0) return responseJson({ success: false, message: 'จำนวนต้องมากกว่า 0' });

      // อัปเดตยอดคงเหลือในตาราง products
      const productsSheet = getSheet(SHEET_NAMES.PRODUCTS);
      const productValues = productsSheet.getDataRange().getValues();
      const headers = productValues[0];
      const idIdx = headers.indexOf('product_id');
      const qtyIdx = headers.indexOf('qty');
      const minIdx = headers.indexOf('min_stock');
      const statusIdx = headers.indexOf('status');
      const updatedIdx = headers.indexOf('updated_at');

      let found = false;
      let newQty = 0;

      for (let i = 1; i < productValues.length; i++) {
        if (String(productValues[i][idIdx]) === pid) {
          found = true;
          const currentQty = Number(productValues[i][qtyIdx] || 0);
          const minStock = Number(productValues[i][minIdx] || 5);

          if (moveType === 'OUT') {
            if (qtyDelta > currentQty) {
              return responseJson({
                success: false,
                message: `จำนวนคงเหลือไม่พอสำหรับการเบิกจ่าย (คงเหลือ ${currentQty} ${productValues[i][headers.indexOf('unit')]})`
              });
            }
            newQty = currentQty - qtyDelta;
          } else {
            // IN หรือ ADJUST บวก
            newQty = currentQty + qtyDelta;
          }

          const newStatus = newQty <= 0 ? 'OUT' : (newQty <= minStock ? 'LOW' : 'OK');
          productsSheet.getRange(i + 1, qtyIdx + 1).setValue(newQty);
          productsSheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
          productsSheet.getRange(i + 1, updatedIdx + 1).setValue(now);
          break;
        }
      }

      if (!found) {
        return responseJson({ success: false, message: 'ไม่พบสินค้านี้ในฐานข้อมูล' });
      }

      // บันทึกลง stock_movements
      const moveRow = [
        movementId,
        pid,
        moveType,
        qtyDelta,
        payload.reason || (moveType === 'IN' ? 'รับเข้า' : 'เบิกจ่าย'),
        payload.note || '',
        payload.created_by || 'admin',
        now
      ];
      movementSheet.appendRow(moveRow);

      return responseJson({
        success: true,
        movementId,
        newQty,
        created_at: now,
        message: 'บันทึกรายการความเคลื่อนไหวสำเร็จ'
      });
    }

    // 5) ปรับยอดสต๊อกโดยตรง (Adjust Stock)
    if (type === 'adjustStock') {
      const pid = String(payload.product_id || payload.pid || '');
      const delta = Number(payload.delta || 0);
      const note = payload.note || 'ปรับยอดจากการตรวจนับสินค้า';

      if (!pid) return responseJson({ success: false, message: 'กรุณาระบุรหัสสินค้า' });

      const productsSheet = getSheet(SHEET_NAMES.PRODUCTS);
      const values = productsSheet.getDataRange().getValues();
      const headers = values[0];
      const idIdx = headers.indexOf('product_id');
      const qtyIdx = headers.indexOf('qty');
      const minIdx = headers.indexOf('min_stock');
      const statusIdx = headers.indexOf('status');
      const updatedIdx = headers.indexOf('updated_at');

      for (let i = 1; i < values.length; i++) {
        if (String(values[i][idIdx]) === pid) {
          const currentQty = Number(values[i][qtyIdx] || 0);
          const minStock = Number(values[i][minIdx] || 5);
          const nextQty = currentQty + delta;
          if (nextQty < 0) {
            return responseJson({ success: false, message: 'ยอดสต๊อกรวมไม่สามารถติดลบได้' });
          }

          const now = new Date().toISOString();
          const newStatus = nextQty <= 0 ? 'OUT' : (nextQty <= minStock ? 'LOW' : 'OK');
          productsSheet.getRange(i + 1, qtyIdx + 1).setValue(nextQty);
          productsSheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
          productsSheet.getRange(i + 1, updatedIdx + 1).setValue(now);

          // บันทึกความเคลื่อนไหว
          const movementSheet = getSheet(SHEET_NAMES.MOVEMENTS);
          movementSheet.appendRow([
            generateId('M'),
            pid,
            delta >= 0 ? 'IN' : 'OUT',
            Math.abs(delta),
            'ปรับยอดสต๊อก',
            note,
            payload.created_by || 'admin',
            now
          ]);

          return responseJson({ success: true, newQty: nextQty, message: 'ปรับปรุงยอดสต๊อกสำเร็จ' });
        }
      }

      return responseJson({ success: false, message: 'ไม่พบรายการสินค้า' });
    }

    return responseJson({ success: false, message: `การทำงานไม่ถูกต้อง: ${type}` });
  } catch (err) {
    return responseJson({
      success: false,
      message: err.message || 'เกิดข้อผิดพลาดในการประมวลผลบนเซิร์ฟเวอร์'
    });
  }
}

function testScript() {
  ensureHeaders();
  console.log('Ensure headers completed successfully.');
}
