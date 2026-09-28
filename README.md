# 2027 THE Rankings Analysis – Netlify Edition

This version is designed to run fully on Netlify.

## Architecture

Frontend: static HTML, CSS and JavaScript in `public/`

Backend: Netlify Function in `netlify/functions/api.mjs`

Excel processing: ExcelJS

Persistent workbook storage: Netlify Blobs

The original workbook in `data/2027 THE Rankings analysis.xlsx` is used to initialize the Blob the first time the deployed application is opened. After that, cell edits are written to the Blob and survive future deploys.

## Deploy with GitHub

1. Create a new GitHub repository.
2. Upload the contents of this project folder to the repository root.
3. In Netlify, choose **Add new project** → **Import an existing project**.
4. Connect GitHub and select the repository.
5. Netlify reads `netlify.toml` automatically.
6. Deploy the site.
7. Open the generated Netlify URL. The first API request initializes the Excel workbook in Netlify Blobs.

No database URL or Blob credentials are required when the function runs on the same Netlify site.

## Local test

```bash
npm install
npx netlify dev
```

Open the local URL shown by Netlify CLI.

## Important

Do not use Netlify's simple drag-and-drop static deploy for this project because the application requires serverless Functions and npm dependencies. Git-based deployment is the easiest option.

Formula cells are detected from Excel and shown as calculated/read-only fields. ExcelJS preserves formulas and their stored results but does not itself calculate arbitrary Excel formulas. The supplied workbook currently contains no formula cells.
