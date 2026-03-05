# Cycle Count PWA — 8 Fixes Design

**Date:** 2026-03-05
**Status:** Approved

## 1. Fix Image Display in App

**Root cause:** `@microsoft.graph.downloadUrl` expires after ~1 hour. Cached local records with stale URLs cause silent `<img>` failures.

**Fix:**
- Add `onerror` handler on `existingPhoto` that hides the image and shows "Photo unavailable" text
- When source is a local record's `_photoLocal` (base64 data URL), use it directly (always works offline)
- When source is a OneDrive download URL from the API, it's fresh and valid — no change needed there

**Files:** `app.js` (lookupExistingPhoto), `index.html` (existingPhoto element)

## 2. Scan Mode Toggle: Bin vs SKU

**UX:** Segmented toggle control with two options: "Scan → SKU" (default) and "Scan → Bin". Placed between the Bin/SKU form fields and the "Start Scanner" button inside the Count Entry card.

**Behavior:**
- Toggle persisted to `localStorage('scanTarget')`, defaults to `'sku'`
- Scanner success callback reads the toggle state and writes to the appropriate field (`$('sku')` or `$('bin')`)
- After writing, calls `validateAndCompute()` as it does now

**Files:** `index.html` (new toggle HTML + CSS), `app.js` (scanner callback, localStorage persistence)

## 3. Photo Hyperlink in Excel (Column G)

**Approach:** Write `=HYPERLINK("url", "📷 View Photo")` formula instead of plain URL string.

**Change:** In `addRowToExcel()`, if `row.PhotoUrl` is non-empty, write the HYPERLINK formula to the PhotoUrl cell position. If empty, write empty string as before.

**Files:** `app.js` (addRowToExcel)

## 4. LastCountDate Format + LastCountedBy

**Changes:**
- Format timestamp as `MM-DD-YYYY hh:mm AM/PM` before writing to OnHand's LastCountDate column
- Add new `LastCountedBy` column to OnHand table (via `ensureOnHandColumns`)
- Populate `LastCountedBy` with `account.name` in `updateOnHandRow()`
- Update inventory cache structure to include `lastCountedBy`
- Update the "Last counted" info display: "Last counted: **10 units** · 50m ago by Joshua Taitt"

**Files:** `app.js` (ensureOnHandColumns, updateOnHandRow, loadInventoryCache, validateAndCompute)

## 5. Fix Header Clipping

**Root cause:** `.wrap` has `padding: 16px` which doesn't account for mobile safe-area insets. The h1 and theme toggle can be partially covered by the status bar or become untappable.

**Fix:**
- Change `.wrap` padding-top to `max(16px, env(safe-area-inset-top))`
- Ensure `.theme-toggle` meets 44x44px minimum tap target
- Add `z-index: 10` to `.header-top` to prevent overlap issues

**Files:** `index.html` (CSS only)

## 6. Background KPI Sync

**Change:** Replace `await refreshKPIs()` with fire-and-forget pattern:

```javascript
refreshKPIs()
  .then(() => toast('KPI sheet updated', 'success'))
  .catch(e => console.warn('KPI refresh failed:', e.message))
```

The sync button re-enables immediately after records finish writing. KPI generation runs silently in the background.

**Files:** `app.js` (syncNow)

## 7. Confirm Modal — Dynamic Button Text

**Change:** Extend `showConfirm(title, msg, okText = 'OK')` to accept a third parameter for the confirm button label.

- Update the modal's OK button text dynamically
- Update callers: Add Inventory → "Confirm", Clear Records → "Delete"

**Files:** `index.html` (change default button text from "Delete" to "OK"), `app.js` (showConfirm signature, callers)

## 8. Scanner Auto-Close After Scan

**Change:** In the scanner success callback, call `stopScan()` immediately after populating the field. One scan → populate → close camera.

```javascript
text => {
  const target = scanTarget === 'bin' ? 'bin' : 'sku'
  $(target).value = text.trim()
  toast('Scanned: ' + text.trim(), 'success')
  validateAndCompute()
  stopScan()  // auto-close
}
```

**Files:** `app.js` (startScan callback)

## Service Worker

Bump cache version to `cyclecount-v9` after all changes.
