# Web app (Netlify)

This `web` branch is a browser version of CSV → Excel Generator with the same Shopify validation, Excel report, corrected XLSX, SQL export, and Variation/Product template flows.

## Deploy on Netlify

1. In Netlify: **Add new site → Import from Git** → this repository.
2. Set **branch** to `web`.
3. Build settings (also in `netlify.toml`):
   - Base directory: `web`
   - Build command: `npm run build`
   - Publish directory: `web/dist` (Netlify uses `publish = "dist"` relative to base)

## Local

```bash
cd web
npm install
npm run dev
```

## Numbers files

The desktop app can convert `.numbers` via Aspose or Apple Numbers. In the browser, conversion only works if the package already contains a CSV/TSV. Otherwise export CSV from Numbers first.
