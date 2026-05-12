# Flex Email Generator (Excel) — Browser Extension

A Chrome / Edge (Manifest V3) port of the original
[`FlexEmailGeneratorExcelBased.user.js`](../FlexEmailGeneratorExcelBased.user.js)
Tampermonkey userscript.

It injects a floating panel into UL Project Details pages that
generates emails (plain `mailto:`, scheduled, external-link, and the
Travel Approval flow) from templates **stored in a local `.xlsx`
file** — typically a workbook your team co-edits on OneDrive and then
syncs back to your machine. The extension reads the file directly
through the browser's File System Access API; no template data ever
leaves your computer.

---

## Features

- 🧩 Same UX as the userscript — floating "✉ Email Generator" panel on
  `https://portal.ul.com/Project/Details/*`.
- 📂 Bind a local `.xlsx` once, then click **🔄 Reload** to pull the
  latest version your colleagues edited on OneDrive.
- 🧾 Default mode ships with a built-in YAML template set, identical
  to the original userscript.
- 🎛 Toolbar popup mirrors the old Tampermonkey menu commands
  (Toggle panel · Download example · Pick · Reload · Forget).
- 🔒 100% local — no remote `@require`, no analytics, no host
  permissions outside `portal.ul.com`. The js-yaml and SheetJS
  libraries are bundled inside the extension.

---

## Install (developer / unpacked)

1. **Download or clone** this repository.
2. Open Chrome / Edge and go to **`chrome://extensions`**
   (Edge: `edge://extensions`).
3. Enable **Developer mode** (toggle in the top-right corner).
4. Click **Load unpacked** and select the
   [`extension/`](.) folder of this repo.
5. Pin **Flex Email Generator** to your toolbar (puzzle-piece icon →
   📌). Open any
   `https://portal.ul.com/Project/Details/...` page and you should see
   the floating ✉ panel appear automatically.

> The first time you click **📂 Pick Excel**, the browser will ask
> you to select your `.xlsx` template file. Pick the OneDrive-synced
> copy on your local disk — its handle is persisted in the
> extension's IndexedDB so subsequent reloads don't need re-picking.

---

## How to use

| Action | Where |
|---|---|
| Toggle the in-page panel | toolbar popup → **👁 Toggle in-page panel**, or hover the small ✉ tab on the page |
| Download an example workbook to start customizing | toolbar popup → **⬇️ Download example .xlsx**, **or** in-page panel button |
| Bind a local `.xlsx` template | in-page panel → **📂 Pick Excel** *(recommended; popup may also work)* |
| Re-read the latest content from the bound file | in-page panel → **🔄 Reload**, or popup → **🔄 Reload from Excel** |
| Forget the bound file | popup → **🗑 Forget Excel binding** |

> **Note on Pick from popup.** Some browsers require the file picker
> to be triggered from a click on the page itself rather than an
> extension popup. If clicking **📂 Pick Excel…** in the popup does
> nothing, just open the in-page panel and use its **📂 Pick Excel**
> button instead — it always works.

For the full Excel sheet schema (column names, supported types,
variables, etc.) see the long comment block at the top of
[`FlexEmailGeneratorExcelBased.user.js`](../FlexEmailGeneratorExcelBased.user.js).
The extension uses exactly the same parser and template format.

---

## Folder layout

```
extension/
├── manifest.json            # MV3 manifest
├── content/
│   ├── gm-compat.js         # tiny GM_* / unsafeWindow shim
│   └── main.js              # the original userscript body
├── lib/
│   ├── js-yaml.min.js       # js-yaml 4.1.0 (MIT)
│   └── xlsx.full.min.js     # SheetJS 0.18.5 (Apache-2.0)
├── popup/
│   ├── popup.html           # toolbar popup UI
│   ├── popup.css
│   └── popup.js             # forwards commands to content script
├── icons/                   # 16/48/128 PNG icons
└── README.md                # this file
```

### Differences vs. the userscript

| Userscript | Extension equivalent |
|---|---|
| `@require` js-yaml / xlsx CDN | Bundled in `extension/lib/`, loaded via `content_scripts` |
| `GM_getValue` / `GM_setValue` | `chrome.storage.local` (sync-mirrored cache, see `gm-compat.js`) |
| `GM_setClipboard` | `navigator.clipboard.writeText` |
| `GM_registerMenuCommand` | Toolbar popup buttons → `chrome.runtime.sendMessage` |
| `unsafeWindow` | `window` (content scripts can call `window.showOpenFilePicker` directly) |

`content/main.js` is otherwise the **same code** as
`FlexEmailGeneratorExcelBased.user.js` (with the userscript header
removed and the menu-registration block replaced by a message
listener), so any future update to the userscript can be re-applied
with a small diff.

---

## Build / packaging

There is no build step — the extension is plain HTML/CSS/JS.

To produce a `.zip` you can upload to the Chrome Web Store / Edge
Add-ons:

```bash
cd extension
zip -r ../flex-email-generator.zip . -x '*.DS_Store'
```

---

## Browser support

- Chrome 100+ and Chromium-based Edge 100+ on desktop.
- Requires the **File System Access API** (`window.showOpenFilePicker`),
  which is available in those browsers.

Firefox is not currently supported — it lacks `showOpenFilePicker`
and uses a different extension manifest schema.

---

## License

The original userscript code in `content/main.js` follows this
repository's license. Bundled third-party libraries keep their own
licenses:

- **js-yaml** 4.1.0 — MIT
- **SheetJS** (xlsx) 0.18.5 — Apache-2.0
