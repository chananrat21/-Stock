/**
 * Google Sheet API Client for Stock Management System
 * Handles communication with Google Apps Script Web App without CORS preflight issues.
 */
(function (global) {
  'use strict';

  const STORAGE_KEY_API_URL = 'stock_api_url';

  function getStoredApiUrl() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY_API_URL);
      if (stored && stored.trim()) return stored.trim();
    } catch (e) {
      // localStorage disabled or not accessible
    }
    return (global.STOCK_API_URL || '').trim();
  }

  function setStoredApiUrl(url) {
    try {
      if (url && url.trim()) {
        localStorage.setItem(STORAGE_KEY_API_URL, url.trim());
      } else {
        localStorage.removeItem(STORAGE_KEY_API_URL);
      }
    } catch (e) {}
  }

  function parseScriptId(url) {
    if (!url) return '';
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/\/macros\/s\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : '';
  }

  function isConfigured(apiUrl) {
    const url = (apiUrl || getStoredApiUrl() || '').trim();
    return !!url && 
      !url.includes('YOUR_DEPLOYMENT_ID') && 
      !url.includes('PASTE_') && 
      !url.includes('AKfycbyYOUR_') &&
      url.startsWith('https://script.google.com/') &&
      url.includes('/exec');
  }

  /**
   * Send POST request to Google Apps Script.
   * NOTE: We intentionally use 'text/plain;charset=utf-8' to avoid CORS preflight (OPTIONS)
   * which Google Apps Script Web Apps do not natively handle.
   */
  async function postRequest(payload, apiUrl) {
    const targetUrl = (apiUrl || getStoredApiUrl() || '').trim();
    if (!targetUrl) {
      throw new Error('ยังไม่ได้กำหนด URL ของ Google Apps Script');
    }
    if (targetUrl.includes('/macros/library/') || !targetUrl.includes('/exec')) {
      throw new Error('URL ไม่ถูกต้อง ต้องเป็น Web App URL ที่ลงท้ายด้วย /exec (สร้างจาก Deploy > New deployment > Web app)');
    }

    const response = await fetch(targetUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });

    const text = await response.text();
    let result;
    try {
      result = JSON.parse(text);
    } catch (e) {
      throw new Error(`เซิร์ฟเวอร์ตอบกลับไม่ถูกต้อง (HTTP ${response.status}): ${text.slice(0, 100)}`);
    }

    if (!response.ok) {
      throw new Error(`HTTP Error ${response.status}: ${result && result.message ? result.message : 'เกิดข้อผิดพลาด'}`);
    }

    return result;
  }

  /**
   * Send GET request to Google Apps Script
   */
  async function getRequest(params = {}, apiUrl) {
    const baseUrl = (apiUrl || getStoredApiUrl() || '').trim();
    if (!baseUrl) {
      throw new Error('ยังไม่ได้กำหนด URL ของ Google Apps Script');
    }
    if (baseUrl.includes('/macros/library/') || !baseUrl.includes('/exec')) {
      throw new Error('URL ไม่ถูกต้อง ต้องเป็น Web App URL ที่ลงท้ายด้วย /exec');
    }

    const query = new URLSearchParams(params).toString();
    const targetUrl = query ? `${baseUrl}${baseUrl.includes('?') ? '&' : '?'}${query}` : baseUrl;

    const response = await fetch(targetUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`HTTP Error ${response.status}`);
    }

    return await response.json();
  }

  // API Methods
  async function testConnection(apiUrl) {
    const url = (apiUrl || getStoredApiUrl() || '').trim();
    if (!url) {
      return { success: false, message: 'กรุณากรอก URL ก่อนทดสอบ' };
    }

    if (url.includes('/macros/library/') || url.includes('/edit')) {
      const scriptId = parseScriptId(url);
      return {
        success: false,
        isLibraryUrl: true,
        scriptId: scriptId,
        editorUrl: `https://script.google.com/d/${scriptId}/edit`,
        message: 'URL ที่คุณระบุเป็นลิงก์ "Library / Editor" ของ Apps Script ไม่ใช่ลิงก์ Web App ที่เรียกใช้งานได้'
      };
    }

    if (!url.includes('/exec')) {
      return {
        success: false,
        message: 'URL ต้องเป็น Web App URL ที่ลงท้ายด้วย /exec (สร้างจาก Deploy > New deployment > Web app)'
      };
    }

    try {
      const res = await getRequest({ type: 'ping' }, url);
      return {
        success: !!(res && (res.success || Array.isArray(res))),
        message: res && res.message ? res.message : 'เชื่อมต่อเซิร์ฟเวอร์สำเร็จ'
      };
    } catch (error) {
      try {
        const prodRes = await getRequest({ type: 'products' }, url);
        if (Array.isArray(prodRes)) {
          return { success: true, message: 'เชื่อมต่อสำเร็จ (พบรายการสินค้า)' };
        }
      } catch (err2) {}
      return { success: false, message: error.message || 'ไม่สามารถเชื่อมต่อได้' };
    }
  }

  async function getProducts(apiUrl) {
    return getRequest({ type: 'products' }, apiUrl);
  }

  async function getMovements(apiUrl) {
    return getRequest({ type: 'stock_movements' }, apiUrl);
  }

  async function getSuppliers(apiUrl) {
    return getRequest({ type: 'suppliers' }, apiUrl);
  }

  async function getDashboard(apiUrl) {
    return getRequest({ type: 'dashboard' }, apiUrl);
  }

  async function addProduct(product, apiUrl) {
    return postRequest({ type: 'addProduct', ...product }, apiUrl);
  }

  async function updateProduct(product, apiUrl) {
    return postRequest({ type: 'updateProduct', ...product }, apiUrl);
  }

  async function deleteProduct(productId, apiUrl) {
    return postRequest({ type: 'deleteProduct', product_id: productId }, apiUrl);
  }

  async function addMovement(movement, apiUrl) {
    return postRequest({ type: 'addMovement', ...movement }, apiUrl);
  }

  async function adjustStock(productId, delta, note = '', apiUrl) {
    return postRequest({ type: 'adjustStock', product_id: productId, delta, note }, apiUrl);
  }

  const StockAPI = {
    getStoredApiUrl,
    setStoredApiUrl,
    parseScriptId,
    isConfigured,
    testConnection,
    getProducts,
    getMovements,
    getSuppliers,
    getDashboard,
    addProduct,
    updateProduct,
    deleteProduct,
    addMovement,
    adjustStock,
    postRequest,
    getRequest
  };

  global.StockAPI = StockAPI;
  global.StockGoogleSheet = StockAPI; // Compatibility
})(typeof window !== 'undefined' ? window : this);
