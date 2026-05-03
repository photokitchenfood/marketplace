# PhotoKitchen — Project Summary

## What This Is
A suite of internal tools for **The Marketplace (TMP)**, a food photography brand based in the Philippines. The goal is to replace a manual Google Slides + Dropbox workflow for browsing, tagging, and watermarking a 10,000+ photo archive. Built as lightweight HTML files — no backend, no database.

---

## The Three Tools

### Tool 1 — Campaign Uploader (`campaign-uploader.html`)
A Claude artifact (lives in this Project as a chat). Used by Andi once per campaign shoot.

- Loads photos from Dropbox (PKCE OAuth) or local folder upload
- Claude's vision API auto-tags each photo: Main Food, Brands, Ingredients, Colors, Styling
- Review grid lets you correct tags before exporting
- Exports a ZIP with two files:
  - `YYMMDD-[campaign name].js` — campaign data with compressed thumbnails (500px, 85% JPEG)
  - `campaigns-index.js` — auto-updated master list of all campaigns
- You push both to GitHub → Catalog website picks them up automatically
- Campaigns folder in repo: `campaigns/` (needs a `.gitkeep` placeholder to exist)

### Tool 2 — Logo Placement Tool (`logo-placement-tool.html`)
A Claude artifact (lives in this Project as a chat). Used by Andi once per campaign for client turnover.

- Loads photos from Dropbox or local upload
- TMP brand logos (Black and White) are embedded as transparent PNGs — no upload needed
- Campaign logo upload with automatic background removal (black-only removal for circular badges)
- Claude's vision API suggests: which corner for TMP logo, which corner for campaign logo, Black or White
- Per-photo override of any suggestion before exporting
- Logo sizes (calibrated to match reference output):
  - TMP logo: 47.5% of image width, 3% padding
  - Campaign badge: 16% × 16%, 3% padding, circular clip
- Exports as ZIP (local download) or uploads directly to Dropbox into a `JPEGS with Logos` folder alongside the source folder at the same level
- Download filenames append `-logo` before the extension

### Tool 3 — The Marketplace Shoot Catalog (`index.html`)
Hosted on GitHub Pages at `https://photokitchenfood.github.io/marketplace/`. Accessed by the whole team.

- Loads campaign data from JS files pushed by the Campaign Uploader
- Sidebar filters: search, campaign dropdown, brand, ingredient, color, year
- Photo grid with cards, tags, and Dropbox links
- **Select mode** — multi-select photos for batch actions
- **Watermark Modal** — for social media use (internal team):
  - Small TMP logo (30% width), Black or White
  - Corner picker (4-position grid)
  - Blend modes: Normal, Multiply, Screen, Soft Light, Overlay
  - Opacity slider (10–100%)
  - Downloads full-res images from Dropbox, composites logo, saves locally
- **Logo Placement flow** — to be built into the Catalog next (for client turnover without needing the Claude artifact):
  - Same corner picker, TMP logo + campaign logo upload
  - Outputs to Dropbox `JPEGS with Logos` folder
  - No AI suggestion (manual placement only, since the Claude API can't be safely embedded in a public GitHub Pages file)

---

## Repository & Infrastructure

- **GitHub repo:** `https://github.com/photokitchenfood/marketplace`
- **GitHub Pages URL:** `https://photokitchenfood.github.io/marketplace/`
- **Dropbox app:** PKCE OAuth flow, redirect URI set to the GitHub Pages URL
- **Dropbox token:** Currently short-lived (PKCE limitation). Option to hardcode a personal access token from the Dropbox developer console for a permanent connection — one-line change in the HTML
- **GitHub Desktop:** Used to push files. After Campaign Uploader exports, workflow is: move JS files into repo folder → commit → push origin

---

## Embedded Assets

The TMP logos are embedded directly in the tool files as base64 transparent PNGs (gray background removed):

- White logo: white text/border on transparent — for dark backgrounds
- Black logo: black text/border on transparent — for light backgrounds
- Both derived from `The_Marketplace_Logo_Vector_WHITE_copy.jpg` and `The_Marketplace_Logo_Vector_BLACK_copy.jpg` with the #999999 gray background stripped

---

## Workflow End-to-End (Post-Shoot)

1. Open **Campaign Uploader** in Claude → load photos from Dropbox JPEG folder or local upload
2. Claude auto-tags all photos → review and correct → export ZIP
3. Move JS files into GitHub repo via GitHub Desktop → commit → push
4. **Catalog website** auto-loads new campaign in the dropdown
5. Team browses photos, uses **Watermark Modal** to download social media versions
6. Andi opens **Logo Placement Tool** in Claude → loads same photos → approves placements → uploads to Dropbox `JPEGS with Logos` for client turnover

---

## What's Built and Working

- ✅ Campaign Uploader — complete, ready to test
- ✅ Logo Placement Tool — complete, sizes calibrated, ZIP export working, campaign logo background removal fixed
- ✅ Watermark Modal — updated with blend modes and opacity slider
- ✅ Catalog core — filtering, search, select mode, card grid

---

## Next Steps (In Order)

1. **Test Campaign Uploader** — upload photos from a real shoot, verify tags, check JS file output and GitHub push workflow
2. **Update Catalog to load from JS files** — replace Import CSV and Import from Dropbox buttons, wire up `campaigns-index.js` dropdown, load campaign data dynamically
3. **Add Logo Placement flow to Catalog** — manual corner picker for client turnover, Dropbox upload to `JPEGS with Logos`
4. **Hardcode Dropbox personal token** — replace PKCE flow in Catalog with a static token from the Dropbox developer console for permanent connection
5. **Final QA** — test full end-to-end workflow with real campaign data
