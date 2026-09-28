(function () {
  const DEFAULT_API = 'https://script.google.com/macros/s/AKfycbyYOUR_DEPLOYMENT_ID/exec';

  function request(payload, apiUrl = DEFAULT_API) {
    return fetch(apiUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    }).then(async (res) => {
      const text = await res.text();
      try {
        return JSON.parse(text);
      } catch (error) {
        return { success: false, message: text };
      }
    });
  }

  function getProducts(apiUrl = DEFAULT_API) {
    return fetch(`${apiUrl}?type=products`)
      .then(res => res.json())
      .catch(() => ({ success: false, message: 'Failed to load products' }));
  }

  function getDashboard(apiUrl = DEFAULT_API) {
    return fetch(`${apiUrl}?type=dashboard`)
      .then(res => res.json())
      .catch(() => ({ success: false, message: 'Failed to load dashboard' }));
  }

  function addProduct(product, apiUrl = DEFAULT_API) {
    return request({ type: 'addProduct', ...product }, apiUrl);
  }

  function updateProduct(product, apiUrl = DEFAULT_API) {
    return request({ type: 'updateProduct', ...product }, apiUrl);
  }

  function addMovement(movement, apiUrl = DEFAULT_API) {
    return request({ type: 'addMovement', ...movement }, apiUrl);
  }

  window.StockGoogleSheet = {
    DEFAULT_API,
    request,
    getProducts,
    getDashboard,
    addProduct,
    updateProduct,
    addMovement
  };
})();
