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

---

---

# 📖 安裝與使用教學 · Installation & Usage Guide

> 本教學適用於一般使用者。技術細節請見上方的 technical 區段。  
> This guide is for end users. For technical details, see the sections above.

---

## 一、這個工具是什麼？ · What is This Extension?

**Flex Email Generator (Excel)** 是一個 Chrome / Edge 瀏覽器擴充功能（Extension），  
專為在 **UL Solutions Portal**（`portal.ul.com`）的 **Project Details** 頁面上快速產生制式電子郵件而設計。

**Flex Email Generator (Excel)** is a Chrome / Edge browser extension  
designed for quickly generating standard emails on the **Project Details** pages of the **UL Solutions Portal** (`portal.ul.com`).

### 解決什麼問題？ · What problems does it solve?

| 問題 Problem | 解決方式 Solution |
|---|---|
| 每次寄信都要手動複製貼上固定格式 · Copying and pasting fixed email formats repeatedly | 從 Excel 模板一鍵生成 · One-click email generation from an Excel template |
| 郵件格式各人不同，難以統一 · Inconsistent email formats across team members | 全團隊共用同一份 OneDrive 上的 `.xlsx` · Share one `.xlsx` template file via OneDrive |
| 需要同時產生多種不同類型的郵件 · Need to generate several email types | 面板上一排按鈕，每個按鈕對應一種模板 · One button per template in the panel |
| 不希望資料上傳到外部伺服器 · Don't want data leaving the local machine | 所有讀取均在本機進行，無任何外部請求 · All reads are local; zero external requests |

---

## 二、安裝步驟 · Installation Steps

### 步驟一：取得擴充功能檔案 · Step 1 – Get the Extension Files

1. 前往 GitHub 頁面，點擊 **Code → Download ZIP**，將整個專案壓縮檔下載到本機。  
   Go to the GitHub page, click **Code → Download ZIP** to download the whole project.

2. 將 ZIP 檔解壓縮到一個固定的資料夾（例如 `C:\Tools\ULS-EmailGenerator`）。  
   Unzip the file into a permanent folder (e.g., `C:\Tools\ULS-EmailGenerator`).

   > ⚠️ **注意 Note：** 擴充功能載入後，這個資料夾 **不能移動或刪除**，否則擴充功能會失效。  
   > After loading, do **not** move or delete this folder — the extension will stop working.

---

### 步驟二：在 Chrome 或 Edge 開啟開發者模式 · Step 2 – Enable Developer Mode

#### Chrome

1. 在網址列輸入 **`chrome://extensions`** 並按 Enter。  
   Type **`chrome://extensions`** in the address bar and press Enter.
2. 開啟右上角的 **「開發人員模式」** 開關（Developer mode toggle）。  
   Toggle on **"Developer mode"** in the top-right corner.

#### Edge

1. 在網址列輸入 **`edge://extensions`** 並按 Enter。  
   Type **`edge://extensions`** in the address bar and press Enter.
2. 開啟左下角的 **「開發人員模式」** 開關。  
   Toggle on **"Developer mode"** in the bottom-left corner.

---

### 步驟三：載入擴充功能 · Step 3 – Load the Extension

1. 點擊 **「載入未封裝」**（Load unpacked）按鈕。  
   Click the **"Load unpacked"** button.
2. 在檔案選擇視窗中，瀏覽到剛才解壓縮的資料夾，**選取 `extension/` 子資料夾**，然後按確認。  
   In the file dialog, navigate to the unzipped folder and **select the `extension/` subfolder**, then confirm.
3. 擴充功能清單中應出現 **"Flex Email Generator (Excel)"**，旁邊有一個信封圖示。  
   "**Flex Email Generator (Excel)**" should now appear in the extensions list with an envelope icon.

---

### 步驟四：將擴充功能固定到工具列 · Step 4 – Pin to Toolbar

1. 點擊瀏覽器右上角的 **拼圖圖示** (🧩 Extensions icon)。  
   Click the **puzzle-piece icon** (🧩) in the top-right corner.
2. 找到 **Flex Email Generator**，點擊旁邊的 **📌 圖釘圖示** 將其固定到工具列。  
   Find **Flex Email Generator** and click the **📌 pin icon** next to it to pin it to the toolbar.

安裝完成！你應該在工具列看到一個信封圖示 ✉。  
Installation complete! You should now see an envelope icon ✉ in the toolbar.

---

## 三、使用流程 · How to Use

### 步驟一：準備 Excel 模板 · Step 1 – Prepare the Excel Template

1. 點擊瀏覽器工具列的 ✉ 圖示，開啟 popup 選單。  
   Click the ✉ icon in the toolbar to open the popup menu.
2. 點擊 **⬇️ Download example .xlsx** 下載範例 Excel 模板。  
   Click **⬇️ Download example .xlsx** to download the sample Excel template.
3. 使用 Microsoft Excel 開啟下載的檔案，按照 Sheet 中的格式**填入或修改**你的郵件模板內容，並儲存。  
   Open the file in Microsoft Excel, **fill in or modify** your email templates according to the sheet format, and save.
4. 如果你的團隊使用 OneDrive，可將此檔案放在 OneDrive 同步資料夾中，讓所有人共用同一份模板。  
   If your team uses OneDrive, place the file in an OneDrive-synced folder so everyone shares the same template.

---

### 步驟二：前往 Project Details 頁面 · Step 2 – Navigate to a Project Details Page

在瀏覽器中登入 UL Solutions Portal，並開啟任何一個 Project Details 頁面：  
Log in to the UL Solutions Portal and open any Project Details page:

```
https://portal.ul.com/Project/Details/...
```

頁面載入完成後，你應該會看到畫面右側出現一個浮動的 **✉ Email Generator** 分頁標籤。  
After the page loads, you should see a floating **✉ Email Generator** tab on the right side of the page.

---

### 步驟三：開啟面板並綁定 Excel · Step 3 – Open the Panel and Bind Your Excel File

1. 點擊 **✉ Email Generator** 浮動標籤，或點擊工具列 popup 的 **👁 Toggle in-page panel** 按鈕，展開面板。  
   Click the **✉ Email Generator** floating tab, or click **👁 Toggle in-page panel** in the toolbar popup to expand the panel.
2. 在面板中點擊 **📂 Pick Excel** 按鈕。  
   Click the **📂 Pick Excel** button in the panel.
3. 在跳出的檔案選擇視窗中，選取你的 `.xlsx` 模板檔案，然後按 **開啟 / Open**。  
   In the file picker dialog, select your `.xlsx` template file and click **Open**.
4. 瀏覽器可能會詢問是否允許存取此檔案，請點擊 **「允許 / Allow」**。  
   The browser may ask for permission to access the file — click **"Allow"**.

面板上的按鈕會更新為你的 Excel 模板中定義的郵件類型。  
The panel buttons will update to reflect the email types defined in your Excel template.

---

### 步驟四：產生郵件 · Step 4 – Generate an Email

1. 在面板中點擊任一**郵件模板按鈕**（例如 "Initial Contact"、"Follow-up" 等）。  
   In the panel, click any **email template button** (e.g., "Initial Contact", "Follow-up", etc.).
2. 工具會根據目前頁面的專案資訊（Project Number、Project Manager 等）自動填入郵件內容，並開啟預覽視窗。  
   The tool will auto-fill email content based on the current page's project info (Project Number, Project Manager, etc.) and open a preview window.
3. 確認郵件內容後，點擊 **✉️ Generate Email** 按鈕，即可透過 `mailto:` 連結開啟預設郵件用戶端。  
   After reviewing, click **✉️ Generate Email** to open your default email client with the `mailto:` link.
4. 如需複製郵件內容，點擊 **📋 Copy** 按鈕將主旨與內文複製到剪貼簿。  
   To copy the content, click **📋 Copy** to copy the subject and body to the clipboard.

---

### 步驟五：日後更新模板 · Step 5 – Update the Template Later

當你或團隊成員修改了 OneDrive 上的 `.xlsx` 模板並同步到本機後：  
When you or a team member has updated the `.xlsx` on OneDrive and it has synced locally:

- 點擊面板的 **🔄 Reload** 按鈕，或  
  Click **🔄 Reload** in the panel, or
- 點擊工具列 popup 的 **🔄 Reload from Excel** 按鈕，  
  Click **🔄 Reload from Excel** in the toolbar popup

即可立即重新讀取最新版本，無需重新選取檔案。  
to immediately re-read the latest version without re-picking the file.

---

## 四、按鈕功能說明 · Button Reference

### 工具列 Popup 按鈕 · Toolbar Popup Buttons

點擊瀏覽器工具列中的 ✉ 圖示即可開啟此 popup。  
Click the ✉ icon in the browser toolbar to open this popup.

| 按鈕 Button | 功能說明 Description |
|---|---|
| 👁 **Toggle in-page panel** | 顯示或隱藏頁面內的浮動 Email Generator 面板。 · Show or hide the floating Email Generator panel on the current page. |
| ⬇️ **Download example .xlsx** | 下載一份範例 Excel 模板檔案，可作為自訂模板的起點。 · Download a sample Excel template file to use as a starting point. |
| 📂 **Pick Excel file…** | 開啟本機檔案選擇器，讓你選取 `.xlsx` 模板並與擴充功能綁定。 · Open a local file picker to select and bind your `.xlsx` template. |
| 🔄 **Reload from Excel** | 重新讀取已綁定的 `.xlsx` 模板，載入最新內容（不需重新選取檔案）。 · Re-read the bound `.xlsx` to load the latest content without re-selecting the file. |
| 🗑 **Forget Excel binding** | 清除已綁定的 Excel 檔案記錄，恢復使用內建 YAML 模板。 · Clear the stored Excel file binding and fall back to the built-in YAML templates. |

> **提示 Tip：** 部分瀏覽器要求檔案選擇器必須從頁面本身觸發，而非從 popup 觸發。  
> 若點擊 **📂 Pick Excel file…** 後無反應，請改用**頁面內面板**的 **📂 Pick Excel** 按鈕。  
> Some browsers require the file picker to be triggered from the page itself, not from the popup.  
> If **📂 Pick Excel file…** does nothing, use the **📂 Pick Excel** button inside the in-page panel instead.

---

### 頁面內面板按鈕 · In-page Panel Buttons

這個浮動面板出現在每個 `portal.ul.com/Project/Details/*` 頁面右側。  
This floating panel appears on every `portal.ul.com/Project/Details/*` page.

| 按鈕 Button | 功能說明 Description |
|---|---|
| 📂 **Pick Excel** | 選取本機 `.xlsx` 模板，效果與 popup 的同名按鈕相同，但**建議優先使用面板版本**（更可靠）。 · Select a local `.xlsx` template. Same as the popup button, but **recommended** for better reliability. |
| 🔄 **Reload** | 重新讀取已綁定的 Excel 模板，取得最新版本。 · Re-read the bound Excel template for the latest version. |
| ⬇️ **Download example** | 下載範例 `.xlsx`，與 popup 按鈕功能相同。 · Download the sample `.xlsx`, same as the popup button. |
| **[郵件模板按鈕 · Email template buttons]** | 每個按鈕對應 Excel 中定義的一種郵件模板，點擊後開啟該模板的預覽與生成介面。 · Each button corresponds to one email template defined in the Excel file. Click to preview and generate that email. |

---

## 五、常見問題 FAQ

### Q1：安裝後在頁面上看不到 ✉ Email Generator 面板？  
### Q1: I installed the extension but don't see the ✉ Email Generator panel on the page?

**中文解答：**  
請確認你開啟的頁面網址是否符合以下格式：  
`https://portal.ul.com/Project/Details/...`  
擴充功能只會在這個特定路徑的頁面上注入面板。如果網址不符，面板不會出現。

**English answer:**  
Make sure the page URL matches the pattern:  
`https://portal.ul.com/Project/Details/...`  
The extension only injects the panel on pages matching that specific path. It won't appear on other pages.

---

### Q2：點擊 📂 Pick Excel 後什麼都沒發生？  
### Q2: Nothing happens when I click 📂 Pick Excel?

**中文解答：**  
這是部分瀏覽器的安全限制——從 popup 觸發的檔案選擇器可能被封鎖。  
解決方法：請直接使用**頁面內浮動面板**的 **📂 Pick Excel** 按鈕（從頁面觸發的選擇器一定有效）。

**English answer:**  
This is a browser security restriction — file pickers triggered from a popup may be blocked.  
**Solution:** Use the **📂 Pick Excel** button inside the **in-page floating panel** instead. This always works because it originates from the page itself.

---

### Q3：瀏覽器彈出「允許存取檔案」的提示，我需要按什麼？  
### Q3: The browser shows a "file access permission" prompt — what should I click?

**中文解答：**  
請點擊 **「允許 / Allow」**（或「Edit files」）。這個提示是瀏覽器的 File System Access API 安全機制，  
讓你明確授權擴充功能讀取指定的本機 Excel 檔案。  
你的檔案 **不會上傳到任何伺服器**，所有讀取都在本機進行。

**English answer:**  
Click **"Allow"** (or "Edit files"). This prompt is the browser's File System Access API security dialog  
asking you to explicitly grant the extension read access to the specified local Excel file.  
Your file is **never uploaded** to any server — all reading happens locally on your machine.

---

### Q4：這個工具支援 Firefox 嗎？  
### Q4: Does this extension support Firefox?

**中文解答：**  
**不支援**。Firefox 目前尚不支援 `window.showOpenFilePicker` (File System Access API)，  
且 Firefox 擴充功能使用不同的 manifest 格式 (也不支援 MV3 的所有功能)。  
請使用 Chrome 100+ 或 Chromium-based Edge 100+。

**English answer:**  
**No.** Firefox currently lacks `window.showOpenFilePicker` (File System Access API)  
and uses a different extension manifest format (also without full MV3 support).  
Please use Chrome 100+ or Chromium-based Edge 100+.

---

### Q5：我的 Excel 模板修改後，按鈕沒有更新？  
### Q5: I updated my Excel template but the panel buttons didn't change?

**中文解答：**  
請點擊面板的 **🔄 Reload** 按鈕（或 popup 的 **🔄 Reload from Excel**）。  
擴充功能不會自動偵測檔案變更，需要手動觸發重新載入。

**English answer:**  
Click the **🔄 Reload** button in the panel (or **🔄 Reload from Excel** in the popup).  
The extension does not auto-detect file changes — you need to manually trigger a reload.

---

### Q6：我的郵件資料會上傳到外部伺服器嗎？  
### Q6: Will my email data be uploaded to an external server?

**中文解答：**  
**不會**。這個擴充功能完全在本機運作：  
- Excel 模板透過 File System Access API 直接從本機讀取。  
- 所有資料（模板內容、Project 資訊）僅存於瀏覽器的本機 Storage。  
- 擴充功能不包含任何分析追蹤、遠端 script 或對外網路請求。

**English answer:**  
**No.** This extension operates entirely locally:  
- The Excel template is read directly from your disk via the File System Access API.  
- All data (template content, project info) is stored only in the browser's local storage.  
- The extension contains zero analytics, remote scripts, or outbound network requests.

---

### Q7：移除擴充功能後，之前選取的 Excel 記錄還在嗎？  
### Q7: If I remove the extension, does my saved Excel binding persist?

**中文解答：**  
不在了。擴充功能的所有本機資料（包括已綁定的 Excel 檔案路徑記錄）  
都儲存在瀏覽器的 Extension Storage 中，移除擴充功能時會一併清除。  
Excel 檔案本身不受影響。

**English answer:**  
No. All local data stored by the extension (including the bound Excel file handle)  
lives in the browser's Extension Storage and is cleared when you remove the extension.  
Your actual Excel file is not affected.

---

### Q8：要如何更新擴充功能到新版本？  
### Q8: How do I update the extension to a new version?

**中文解答：**  
1. 從 GitHub 下載最新版本的 ZIP 並解壓縮，**覆蓋**原有資料夾中的檔案。  
2. 前往 `chrome://extensions`（或 `edge://extensions`）。  
3. 找到 Flex Email Generator，點擊 **🔄 重新整理（Reload）** 圖示。  
4. 擴充功能即會載入最新版本。

**English answer:**  
1. Download the latest ZIP from GitHub, unzip, and **overwrite** the files in the existing folder.  
2. Go to `chrome://extensions` (or `edge://extensions`).  
3. Find Flex Email Generator and click the **🔄 Reload** icon.  
4. The extension will load the latest version.

---

*本教學由 Fi5herL 維護 · This guide is maintained by Fi5herL.*
