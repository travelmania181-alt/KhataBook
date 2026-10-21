# KhataBook starter

A browser-based, mobile-friendly income and expense tracker prototype.

## Current starter features
- Add income and expense entries with category, date, and payment method
- Monthly totals and spending category breakdown
- Download monthly income and expense CSV files (Excel-compatible)
- Local browser storage and responsive UI

## Important limitations
This is a starter prototype, not production-ready financial software. Profile/passcode handling is local-only and is **not secure authentication**. Records live in the current browser's local storage, do not sync across devices, and can be lost if browser data is cleared. Do not use for sensitive or business-critical records until real authentication, secure database, backup, and authorization are implemented and tested.

## Hosting
Publish the repository with GitHub Pages from Settings → Pages → Deploy from a branch → `main` → `/(root)`. After publishing, test the HTTPS URL.
