# Budget

A phone app for tracking household spending against monthly limits per category.

- `index.html`, `manifest.json`, `sw.js`, `icon-*.png`: the app, served by GitHub Pages.
- `apps-script/Alerts.gs`: reads BMO alert emails into the Google Sheet (runs in its own Apps Script project).
- `apps-script/Api.gs` and `apps-script/appsscript.json`: the data service attached to the sheet. The app sends a private key with every request.

No transactions or keys are stored in this repository.
