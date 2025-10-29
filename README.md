# Email List Filter Pro (GitHub-ready)

Single-page SaaS for cleaning, segmenting, and comparing email lists — 100% client‑side.

## Local run
1) Extract the ZIP.  
2) Open `index.html` in your browser.  
   - Or start a tiny server for consistent file APIs:
   ```bash
   python -m http.server 8080
   ```
   then go to http://localhost:8080

## Deploy to GitHub Pages
1. Create a repo (e.g., `EmailListFilterPro`).
2. Upload **index.html**, **styles.css**, **app.js**, **README.md**, and **.nojekyll** to the repo root.
3. Push to `main`.
4. In **Settings → Pages**: choose **Deploy from branch**, `main`, `/ (root)`.

Live at:
```
https://<your-username>.github.io/EmailListFilterPro
```

## Notes
- Processes **.xlsx / .xls / .csv**. Prefers “Complete Leads” → “Leads Title” → first sheet.
- Dedupe by **Email**; keeps the first occurrence; removes blanks/invalid emails.
- Segments “Irrelevant” by keywords/domains (.edu/.mil included). Keywords & deny domains are adjustable.
- Downloads a 3-sheet workbook: **Relevant Data**, **Irrelevant Data**, **Removed Duplicates**.
- Part 2 compares **New Processed** vs **Previous** workbooks and exports **New Only** rows (Relevant/Irrelevant).