# Stock Management System

Static stock control web app with Google Sheet integration support.

## Features
- Inventory overview dashboard
- In/out stock movement logging
- Manage products and quantities
- Google Sheet-ready API flow
- GitHub Pages deployment workflow

## Deploy to GitHub Pages
1. Push this repo to GitHub
2. Go to Repository Settings > Pages
3. Source: GitHub Actions
4. The workflow in `.github/workflows/deploy-pages.yml` will deploy automatically on every push to `main`

## Connect to Google Sheet
1. Open a Google Sheet and create tabs:
   - `products`
   - `stock_movements`
   - `suppliers`
   - `settings`
2. Open Extensions > Apps Script
3. Paste the code from `google-apps-script.gs`
4. Deploy as a Web App
5. Copy the generated URL
6. Add this before your app loads:

```html
<script>
  window.STOCK_API_URL = 'https://script.google.com/macros/s/XXXXXXXX/exec';
</script>
```

7. Reload the page. The app will use the Google Sheet data automatically.

## Local preview
```bash
python3 -m http.server 8000
```
Then open:
```text
http://localhost:8000
```

## Template file
Use the CSV template in `GOOGLE_SHEET_TEMPLATE.csv` as a starting point for the sheet structure.
