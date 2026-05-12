
/*
 * ============================================================
 *  Flex Email Generator (Excel) — Excel 範本格式說明
 * ============================================================
 *
 * 設計目標：同事們習慣在 OneDrive 上協作編輯 .xlsx，雲端修改後再
 * 同步回本地。此腳本利用 File System Access API（showOpenFilePicker）
 * 取得本地 .xlsx 的 FileSystemFileHandle，並將該 handle 持久化到
 * IndexedDB，於每次需要載入模板時即時讀取本地檔案內容（即「Reload」），
 * 即可把雲端最新內容拉進 Tampermonkey 面板。
 *
 * Default 模式：使用內建 YAML（與原 Flex Email Generator 相同）。
 * Customize 模式：從使用者選擇的本地 .xlsx 解析模板。
 *
 * ── Excel 結構 ─────────────────────────────────────────────
 * 一個工作簿包含最多四個工作表（sheet）：
 *
 *  ① 工作表名稱：Plain（純文字信件模板）
 *     每一列代表一個 plain 模板。欄位（不分大小寫）：
 *
 *       id                            必填，唯一識別字串
 *       name                          必填，按鈕顯示名稱
 *       order                         選填，整數；面板按鈕排序依據（越小越上面）
 *                                     未填者排在所有有 order 的模板之後，並保持原本相對順序
 *       buttonLabel                   面板按鈕文字
 *       buttonIcon                    面板按鈕 emoji（例：✉️）
 *       buttonColor                   面板按鈕底色（例：#1976d2）
 *       to                            收件人
 *       cc                            副本
 *       subject                       必填，主旨
 *       body                          必填，內文（可 Alt+Enter 換行）
 *
 *  ② 工作表名稱：External（外部連結模板）
 *     每一列代表一個 external 模板。欄位：
 *
 *       id                            必填，唯一識別字串
 *       name                          必填，按鈕顯示名稱
 *       order                         選填，整數；面板按鈕排序依據（同上）
 *       buttonLabel / buttonIcon / buttonColor
 *       externalUrl                   必填，外部網址
 *       externalParams                多行 key=value（每行一組）
 *       externalOpenInNewTab          TRUE/FALSE，預設 TRUE
 *       externalAddFRDate             TRUE/FALSE，加上 FRDate=今天
 *       emailTo / emailCc /           可選，隨附 mailto；
 *       emailSubject / emailBody      若任一不為空則同時開 mailto
 *
 *  ③ 工作表名稱：Scheduled（排程信件模板）
 *     每一列代表一個 scheduled 模板。欄位：
 *
 *       id                            必填，唯一識別字串
 *       name                          必填，按鈕顯示名稱
 *       order                         選填，整數；面板按鈕排序依據（同上）
 *       buttonLabel / buttonIcon / buttonColor
 *       to                            收件人
 *       cc                            副本
 *       subject                       必填，主旨
 *       body                          必填，內文（可 Alt+Enter 換行）
 *       datePickerLabel               日期選擇器標題
 *       datePickerVariable            必填，使用者選到的日期會綁到此變數名
 *       datePickerFormat              ISO / ROC / yyyy-MM-dd / yyyy/MM/dd
 *       datePickerDefaultOffsetDays   整數，預設 = 今天 + N 天
 *       datePickerMin                 today 或 yyyy-MM-dd
 *       datePickerMax                 yyyy-MM-dd 或留空
 *
 *  （向下相容）工作表名稱：Templates
 *     舊版單一工作表格式，仍可繼續使用。欄位需加上 kind 欄
 *     （plain / external / scheduled），其餘欄位同上。
 *     若同時存在 Plain / External / Scheduled 任一工作表，則
 *     優先使用新格式，Templates 工作表將被忽略。
 *
 *  ④ 工作表名稱：TravelApproval（可選；若不存在則使用內建預設）
 *     兩欄：Field | Value（第一列為標頭）。Field 為點分路徑，例：
 *
 *       enabled                            TRUE/FALSE
 *       button.label                       Travel Approval
 *       button.icon                        ✈️
 *       button.color                       #4A6F8A
 *       modal.title                        ✈️ Travel Approval
 *       modal.sections.dates               📅 Travel Dates ...
 *       modal.sections.locations           📍 Travel Locations
 *       modal.sections.from                Departure Location:
 *       modal.sections.to                  Destination Location:
 *       modal.sections.billing             📊 Billing Type
 *       modal.sections.amount              💰 Estimated Travel Cost
 *       modal.sections.preview             📧 Email Preview
 *       modal.placeholders.from            Enter departure location
 *       modal.placeholders.to              Enter destination location
 *       modal.placeholders.amount          Enter estimated amount
 *       modal.labels.noDatesSelected       (No dates selected)
 *       modal.labels.notFilled             (not filled)
 *       options.locations                  多行，每行一個地點
 *       options.billingTypes               多行，每行一種 billing type
 *       emailTemplates.subject             Travel Approval - {{pjNum}} ...
 *       emailTemplates.body                多行內文模板
 *
 * 變數語法與原版相同：{{var}} / {{ var }} / 舊式 #var# 皆支援。
 * 可用變數見 PageData.gather()。
 *
 * 使用流程：
 *   1) 點面板「⬇ Download template」下載範例 .xlsx 開始編輯。
 *   2) 將檔案放到 OneDrive 同步資料夾，由同事在雲端維護。
 *   3) 點面板「📂 Pick Excel」選擇本地檔案；handle 會記到 IndexedDB。
 *   4) 之後需要更新時點「🔄 Reload」即可重新從本地 .xlsx 拉取最新內容。
 *   5) 如要解除綁定，可用「🗑 Forget」清掉 IndexedDB 中的 handle。
 * ============================================================
 */

(function () {
    'use strict';

    // ============================================================
    //  常數
    // ============================================================
    const MODE_KEY = 'fegx.template.mode';
    const TEMPLATE_MODE_DEFAULT = 'default';
    const TEMPLATE_MODE_CUSTOMIZE = 'customize';
    // IndexedDB used for persisting the FileSystemFileHandle of the user-picked .xlsx
    const IDB_NAME = 'fegx-excel-store';
    const IDB_STORE = 'handles';
    const IDB_HANDLE_KEY = 'customizeExcel';
    const TEMPLATE_FILE_NAME = 'flex-email-templates.xlsx';
    const PANEL_ID    = 'feg-panel';
    const PANEL_TRIGGER_TEXT = '✉ Email Generator';
    const PANEL_ARIA_NAME = 'EmailGeneratorArea';
    const PANEL_ACTION_EXPAND = 'Flex';
    const PANEL_ACTION_COLLAPSE = 'Fix';
    const PANEL_HOVER_LEAVE_DELAY_MS = 700;
    let travelLocListIdSeq = 0;

    // ============================================================
    //  內建預設模式 YAML（Default mode）
    // ============================================================
    const DEFAULT_MODE_YAML = `version: 1
templates:
  # === Inactive letter templates ===
  - id: notice-letter
    kind: plain
    name: Notice Inactive Letter
    order: 1
    button:
      label: Notice
      icon: "📬"
      color: "#6B6B6B"
    to: "{{clientEmail}}"
    cc: ""
    subject: "Project Inactive Letter–Project #{{pjNum}}"
    body: |
      Dear Customer,

      Thank you for your trust and support of UL services.

      Regarding the UL certification project submitted by your company on {{dateBooked}}, with Project No. {{pjNum}}, Service Order No. {{odrNum}}, and project description {{pjScope}}, we regret to inform you that the project is unable to proceed with the certification evaluation due to the lack of the required information listed below. As of now, the project status has been changed to inactive.

      In order to proceed with the project, we kindly request your company to provide the following information at your earliest convenience:
      {{projectHoldReason}}

      Please note that the project can only be resumed after UL receives the complete and correct information as requested above. Should you have any questions, please feel free to contact us at any time. We will continue to actively follow up with your company to facilitate the resumption of your certification project.

      Best regards,

      Project Handler / Email:
      {{projectHandlerEmail}}

  - id: inactive-1
    kind: plain
    name: Inactive Letter 1
    order: 2
    button:
      label: "1st Notice"
      icon: "📬"
      color: "#6B6B6B"
    to: "{{clientEmail}}"
    cc: ""
    subject: "The 1st project inactive follow up letter–Project {{pjNum}}"
    body: |
      Dear Customer,

      Thank you for your trust and support of UL services.

      Regarding the UL certification project submitted by your company on {{dateBooked}}, with Project No. {{pjNum}}, Service Order No. {{odrNum}}, and project description {{pjScope}}, we regret to inform you that the project is unable to proceed with the certification evaluation due to the lack of the required information listed below. As of now, the project status has been changed to inactive.

      In order to proceed with the project, we kindly request your company to provide the following information at your earliest convenience:
      {{projectHoldReason}}

      As of today, we have not yet received the complete and correct information from your company, and the project remains inactive.

      Please note that the project can only be resumed after UL receives the complete and correct information as requested above. Should you have any questions, please feel free to contact us at any time. We will continue to actively follow up with your company to facilitate the resumption of your certification project.

      Best regards,

      Project Handler / Email:
      {{projectHandlerEmail}}

  - id: inactive-2
    kind: plain
    name: Inactive Letter 2
    order: 3
    button:
      label: "2nd Notice"
      icon: "📬"
      color: "#4A4A4A"
    to: "{{clientEmail}}"
    cc: ""
    subject: "The 2nd project inactive follow up letter–Project {{pjNum}}"
    body: |
      Dear Customer,

      Thank you for your trust and support of UL services.

      Regarding the UL certification project submitted by your company on {{dateBooked}}, with Project No. {{pjNum}}, Service Order No. {{odrNum}}, and project description {{pjScope}}, we regret to inform you that the project is unable to proceed with the certification evaluation due to the lack of the required information listed below. As of now, the project status has been changed to inactive.

      In order to proceed with the project, we kindly request your company to provide the following information at your earliest convenience:
      {{projectHoldReason}}

      One month ago, we issued the first inactive follow-up notice. However, as of today, we still have not received the complete and correct information from your company, and the project remains inactive.

      Please note that the project can only be resumed after UL receives the complete and correct information as requested above. Should you have any questions, please feel free to contact us at any time. We will continue to actively follow up with your company to facilitate the resumption of your certification project.

      Best regards,

      Project Handler / Email:
      {{projectHandlerEmail}}

  - id: inactive-3
    kind: plain
    name: Inactive Letter 3
    order: 4
    button:
      label: "3rd Notice"
      icon: "📬"
      color: "#333333"
    to: "{{clientEmail}}"
    cc: ""
    subject: "The 3rd project inactive follow up letter–Project {{pjNum}}"
    body: |
      Dear Customer,

      Thank you for your trust and support of UL services.

      Regarding the UL certification project submitted by your company on {{dateBooked}}, with Project No. {{pjNum}}, Service Order No. {{odrNum}}, and project description {{pjScope}}, we regret to inform you that the project is unable to proceed with the certification evaluation due to the lack of the required information listed below. As of now, the project status has been changed to inactive.

      In order to proceed with the project, we kindly request your company to provide the following information at your earliest convenience:
      {{projectHoldReason}}

      Two months ago, we issued the first inactive follow-up notice, and one month ago, we issued the second inactive follow-up notice. However, as of today, we still have not received the complete and correct information from your company, and the project remains inactive.

      Please note that the project can only be resumed after UL receives the complete and correct information as requested above. Should you have any questions, please feel free to contact us at any time. We will continue to actively follow up with your company to facilitate the resumption of your certification project.

      Best regards,

      Project Handler / Email:
      {{projectHandlerEmail}}

  # === Scheduled letter templates ===
  - id: final-notice
    kind: scheduled
    name: Final Notice
    order: 5
    button:
      label: Final Notice
      icon: "⏰"
      color: "#A60F26"
    datePicker:
      label: "Please select the deadline"
      variable: deadlineDate
      format: "ROC"
      defaultOffsetDays: 14
      min: today
      max: ""
    to: "{{clientEmail}}"
    subject: "FINAL NOTICE - {{pjNum}}"
    body: |
      Dear Customer,

      Thank you for your trust and support of UL services.

      Regarding the UL certification project submitted by your company on {{dateBooked}}, with Project No. {{pjNum}}, Service Order No. {{odrNum}}, and project description {{pjScope}}. Approximately four months ago, we informed you in writing that the project was unable to proceed with the certification evaluation due to the lack of the required information listed below, and the project status was changed to inactive.

      In order to proceed with the project, we kindly request your company to provide the following information at your earliest convenience, or provide a clear timeline for submission:
      {{projectHoldReason}}

      As no valid response has been received, approximately three months ago we issued the first inactive follow-up notice, followed by the second and third notices in the subsequent two months. However, as of today, we still have not received the complete and correct information from your company, and the project has remained inactive.

      Based on the project status over the past four months, this letter serves as the final notice. We kindly ask you to take this matter seriously. If within the next two weeks, that is, before {{deadlineDate}}, you are still unable to provide the complete and correct required information and/or samples, we regret to inform you that we will terminate Service Order No. {{odrNum}} and the associated project {{pjNum}}.

      Upon termination, we will charge fees for the services already performed. If payment was made in advance, the remaining balance will be refunded after deducting the applicable charges. Should you wish to resume the service in the future, you may request a new official quotation (valid for three months). We will reassess your service needs and issue a new quotation accordingly.

      Best regards,

      Project Handler / Email:
      {{projectHandlerEmail}}

  # === External report templates ===
  - id: tat-letter
    kind: external
    name: TAT Letter
    order: 6
    button:
      label: TAT Letter
      icon: "🔗"
      color: "#333333"
    external:
      url: "https://epic.ul.com/Report"
      params:
        TemplateUNID: "AHL TAT Letter"
        SelectedOutputType: ".eml"
        ProjectID: "{{projectAnchorHref}}"
        isWorkbench: "False"
      openInNewTab: true
      addFRDate: true

  - id: ecd-letter
    kind: external
    name: ECD Letter
    order: 7
    button:
      label: ECD Letter
      icon: "🔗"
      color: "#333333"
    external:
      url: "https://epic.ul.com/Report"
      params:
        TemplateUNID: "AHL ECD Letter"
        SelectedOutputType: ".eml"
        ProjectID: "{{projectAnchorHref}}"
        isWorkbench: "False"
      openInNewTab: true
      addFRDate: true
    email:
      to: "{{clientEmail}}"
      subject: "ECD Update - {{pjNum}}"
      body: |
        Dear {{clientName}},

        Please refer to the attached ECD report for project {{pjNum}}.

        Regards,
        {{projectHandlerEmail}}

  - id: noa-letter
    kind: external
    name: NOA Letter
    order: 8
    button:
      label: NOA Letter
      icon: "🔗"
      color: "#333333"
    external:
      url: "https://epic.ul.com/Report"
      params:
        TemplateUNID: "Notice of Authorization or Completion Letter"
        SelectedOutputType: ".default"
        ProjectID: "{{projectAnchorHref}}"
        isWorkbench: "False"
      openInNewTab: true
      addFRDate: true

  # === Plain email templates ===
  - id: close-letter
    kind: plain
    name: Close Letter
    order: 10
    button:
      label: Close Letter
      icon: "✉️"
      color: "#333333"
    to: "{{clientEmail}}"
    cc: ""
    subject: "Project Closure - {{pjNum}}"
    body: |
      Dear {{clientName}},

      Project {{pjNum}} - {{pjScope}} has been completed.
      If there are no further questions, please reply “OK” or “Agree” to acknowledge.

      Best regards,
      {{projectHandlerEmail}}

  - id: pi-letter
    kind: external
    name: PI Letter
    order: 9
    button:
      label: PI Letter
      icon: "🔗"
      color: "#333333"
    external:
      url: "https://epic.ul.com/Report"
      params:
        TemplateUNID: "AHL Preliminary Evaluation"
        SelectedOutputType: ".default"
        ProjectID: "{{projectAnchorHref}}"
        isWorkbench: "False"
      openInNewTab: true
      addFRDate: false
`;

    // ============================================================
    //  內建客製模式 YAML（Customize mode 首次執行時使用）
    // ============================================================
    const DEFAULT_YAML = `version: 1
templates:

  # === Inactive Letters 系列 ===
  - id: notice-letter
    kind: plain
    name: Notice Inactive Letter
    order: 1
    button:
      label: Notice (CN)
      icon: "📬"
      color: "#6B6B6B"
    to: "{{clientEmail}}"
    cc: ""
    subject: "Project Inactive Letter–Project #{{pjNum}}"
    body: |
      尊敬的客戶：

      感謝您及貴公司對UL服務的信任與支持。

      關於貴公司{{rocDateBooked}}提交的UL認證項目，認證項目編號 {{pjNum}} ，服務訂單編號 {{odrNum}} ，認證申請描述 {{pjScope}}，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。

      為了繼續推進項目進程，我們需要貴公司盡快提供如下信息
      {{projectHoldReason}}

      請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。

      順祝
      商祺

      Project Handler /Email:
      {{projectHandlerEmail}}

  - id: inactive-1
    kind: plain
    name: Inactive Letter 1
    order: 2
    button:
      label: "1st Notice (CN)"
      icon: "📬"
      color: "#6B6B6B"
    to: "{{clientEmail}}"
    cc: ""
    subject: "The 1st project inactive follow up letter–Project {{pjNum}}"
    body: |
      尊敬的客戶：

      感謝您及貴公司對UL服務的信任與支持。

      關於貴公司{{rocDateBooked}}提交的UL認證項目，認證項目編號 {{pjNum}} ，服務訂單編號 {{odrNum}} ，認證申請描述 {{pjScope}}，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。

      為了繼續推進項目進程，我們需要貴公司盡快提供如下信息
      {{projectHoldReason}}

      到目前為止，我們尚未從貴公司收到完整並正確的上述信息，項目仍處於暫停狀態。

      請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。

      順祝
      商祺

      Project Handler /Email:
      {{projectHandlerEmail}}

  - id: inactive-2
    kind: plain
    name: Inactive Letter 2
    order: 3
    button:
      label: "2nd Notice (CN)"
      icon: "📬"
      color: "#4A4A4A"
    to: "{{clientEmail}}"
    cc: ""
    subject: "The 2nd project inactive follow up letter–Project {{pjNum}}"
    body: |
      尊敬的客戶：

      感謝您及貴公司對UL服務的信任與支持。

      關於貴公司{{rocDateBooked}}提交的UL認證項目，認證項目編號 {{pjNum}} ，服務訂單編號 {{odrNum}} ，認證申請描述 {{pjScope}}，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。

      為了繼續推進項目進程，我們需要貴公司盡快提供如下信息
      {{projectHoldReason}}

      一個月前，我們發出第一次項目暫停跟進通知書，但到目前為止，我們仍未從貴公司收到完整並正確的上述信息，項目仍處於暫停狀態。

      請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。

      順祝
      商祺

      Project Handler /Email:
      {{projectHandlerEmail}}

  - id: inactive-3
    kind: plain
    name: Inactive Letter 3
    order: 4
    button:
      label: "3rd Notice (CN)"
      icon: "📬"
      color: "#333333"
    to: "{{clientEmail}}"
    cc: ""
    subject: "The 3rd project inactive follow up letter–Project {{pjNum}}"
    body: |
      尊敬的客戶：

      感謝您及貴公司對UL服務的信任與支持。

      關於貴公司{{rocDateBooked}}提交的UL認證項目，認證項目編號 {{pjNum}} ，服務訂單編號 {{odrNum}} ，認證申請描述 {{pjScope}}，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。

      為了繼續推進項目進程，我們需要貴公司盡快提供如下信息
      {{projectHoldReason}}

      二個月前，我們發出第一次項目暫停跟進通知書，並且在一個月前，向貴司發出第二次項目暫停跟進通知書，但是到目前為止，我們仍未從貴公司收到完整併正確的上述信息，項目依舊處於暫停狀態。

      請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。

      順祝
      商祺

      Project Handler /Email:
      {{projectHandlerEmail}}

  # === 類型 3：純文字 + 時間選擇面板 ===
  - id: final-notice
    kind: scheduled
    name: Final Notice
    order: 5
    button:
      label: Final Notice (CN)
      icon: "⏰"
      color: "#A60F26"
    datePicker:
      label: "請選擇截止日期 (Deadline)"
      variable: deadlineDate
      format: "ROC"
      defaultOffsetDays: 14
      min: today
      max: ""
    to: "{{clientEmail}}"
    subject: "FINAL NOTICE - {{pjNum}}"
    body: |
      尊敬的客戶：

      感謝您及貴公司對UL服務的信任與支持。

      關於貴公司{{rocDateBooked}}提交的UL認證項目，認證項目編號 {{pjNum}} ，服務訂單編號： {{odrNum}} ，認證申請描述 {{pjScope}}。約四個月前，我們曾書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。項目進度變更為暫停狀態。

      為了繼續推進項目進程，我們需要貴公司盡快提供如下信息，或就該等信息和材料的提交提出明確的時間表。
      {{projectHoldReason}}

      由於未能收到有效反饋，約三個月前，我們發出第一次項目暫停跟進通知書，並且在隨後二個月，接連向貴司發出第二次以及第三次項目暫停跟進通知書。但是到目前為止，我們仍未從貴公司收到完整併正確的上述信息，項目始終處於暫停狀態。

      根據過去四個月的項目進展狀況，我們在此最後一次向您發出項目提醒函，請您務必引起重視，如果該項目在未來兩週内，也就是 {{deadlineDate}} 前您仍不能提供完整併正確的上述必要信息/樣品，我們不得不遺憾的通知您，我們將終止貴公司【{{odrNum}}】號服務訂單及其項下之 {{pjNum}} 認證項目。項目終止後，我們將就我司已經提供的服務向您收取相應的費用；項目採用預付款方式支付的，我們將在扣除必要費用後，退還您的剩餘款項。如以上服務需求在未來需要再次啟動，您可以向我們索取一份新的正式報價（報價有效期為三個月）。我們將重新核定貴公司的服務需求，並向您發出新的報價。

      順祝
      商祺

      Project Handler /Email:
      {{projectHandlerEmail}}

  # === 類型 2：文字 + 開啟外部網頁（TAT Letter） ===
  - id: tat-letter
    kind: external
    name: TAT Letter
    order: 6
    button:
      label: TAT Letter
      icon: "🔗"
      color: "#333333"
    external:
      url: "https://epic.ul.com/Report"
      params:
        TemplateUNID: "AHL TAT Letter"
        SelectedOutputType: ".eml"
        ProjectID: "{{projectAnchorHref}}"
        isWorkbench: "False"
      openInNewTab: true
      addFRDate: true

  # === 類型 2：文字 + 開啟外部網頁（ECD Letter） ===
  - id: ecd-letter
    kind: external
    name: ECD Letter
    order: 7
    button:
      label: ECD Letter
      icon: "🔗"
      color: "#333333"
    external:
      url: "https://epic.ul.com/Report"
      params:
        TemplateUNID: "AHL ECD Letter"
        SelectedOutputType: ".eml"
        ProjectID: "{{projectAnchorHref}}"
        isWorkbench: "False"
      openInNewTab: true
      addFRDate: true
    email:
      to: "{{clientEmail}}"
      subject: "ECD Update - {{pjNum}}"
      body: |
        Dear {{clientName}},

        Please refer to the attached ECD report for project {{pjNum}}.

        Regards,
        {{projectHandlerEmail}}

  # === 類型 2：文字 + 開啟外部網頁（NOA Letter） ===
  - id: noa-letter
    kind: external
    name: NOA Letter
    order: 8
    button:
      label: NOA Letter
      icon: "🔗"
      color: "#333333"
    external:
      url: "https://epic.ul.com/Report"
      params:
        TemplateUNID: "Notice of Authorization or Completion Letter"
        SelectedOutputType: ".default"
        ProjectID: "{{projectAnchorHref}}"
        isWorkbench: "False"
      openInNewTab: true
      addFRDate: true

  # === 類型 1：純文字信件（直接組 mailto 開信） ===
  - id: close-letter
    kind: plain
    name: Close Letter
    order: 10
    button:
      label: Close Letter
      icon: "✉️"
      color: "#333333"
    to: "{{clientEmail}}"
    cc: ""
    subject: "Project Closure - {{pjNum}}"
    body: |
      Dear {{clientName}},

      Project {{pjNum}} - {{pjScope}} has been completed.
      If there are no further questions, please reply “OK” or “Agree” to acknowledge.
      
      Best regards,
      {{projectHandlerEmail}}

  # === 類型 2：文字 + 開啟外部網頁（PI Letter） ===
  - id: pi-letter
    kind: external
    name: PI Letter
    order: 9
    button:
      label: PI Letter
      icon: "🔗"
      color: "#333333"
    external:
      url: "https://epic.ul.com/Report"
      params:
        TemplateUNID: "AHL Preliminary Evaluation"
        SelectedOutputType: ".default"
        ProjectID: "{{projectAnchorHref}}"
        isWorkbench: "False"
      openInNewTab: true
      addFRDate: false

travelApproval:
  enabled: true
  button:
    label: "Travel Approval"
    icon: "✈️"
    color: "#4A6F8A"
  modal:
    title: "✈️ Travel Approval"
    sections:
      dates: "📅 Travel Dates (Multiple Selection)"
      locations: "📍 Travel Locations"
      from: "Departure Location:"
      to: "Destination Location:"
      billing: "📊 Billing Type"
      amount: "💰 Estimated Travel Cost"
      preview: "📧 Email Preview"
    placeholders:
      from: "Enter departure location"
      to: "Enter destination location"
      amount: "Enter estimated amount (e.g. 2000)"
    labels:
      noDatesSelected: "(No dates selected)"
      notFilled: "(not filled)"
  options:
    locations:
      - "ULS Company site"
      - "Customer Site"
    billingTypes:
      - "Non-Billable"
      - "Billable Invoiceable"
      - "Billable Non-Invoiceable"
  emailTemplates:
    subject: "Travel Approval - {{pjNum}} {{datesStr}}"
    body: |
      Dear Manager,

      I would like to request travel approval with the following details:

      Project: {{pjNum}}{{pjNamePart}}
      Description: {{pjScope}}

      Travel Dates: {{datesStr}}
      Departure Location: {{fromLocation}}
      Destination Location: {{toLocation}}

      Billing Type: {{billingType}}
      Estimated Travel Cost: {{amount}}

      Please approve this request.

      Thank you.
      {{projectHandlerEmail}}
`;

    // ============================================================
    //  Storage 模組（GM 儲存：模板模式）
    // ============================================================
    const Storage = {
        /** 讀取模板模式（default / customize） */
        loadTemplateMode() {
            return GM_getValue(MODE_KEY, TEMPLATE_MODE_DEFAULT);
        },
        /** 儲存模板模式（default / customize） */
        saveTemplateMode(mode) {
            GM_setValue(MODE_KEY, mode);
        }
    };

    // ============================================================
    //  ExcelStore 模組（IndexedDB：持久化 FileSystemFileHandle）
    // ============================================================
    const ExcelStore = {
        _dbPromise: null,
        _openDb() {
            if (this._dbPromise) return this._dbPromise;
            this._dbPromise = new Promise((resolve, reject) => {
                if (!('indexedDB' in window)) {
                    reject(new Error('IndexedDB is not available in this browser.'));
                    return;
                }
                const req = window.indexedDB.open(IDB_NAME, 1);
                req.onupgradeneeded = () => {
                    const db = req.result;
                    if (!db.objectStoreNames.contains(IDB_STORE)) {
                        db.createObjectStore(IDB_STORE);
                    }
                };
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error || new Error('Failed to open IndexedDB.'));
            });
            return this._dbPromise;
        },
        async _tx(mode, op) {
            const db = await this._openDb();
            return new Promise((resolve, reject) => {
                const tx = db.transaction(IDB_STORE, mode);
                const store = tx.objectStore(IDB_STORE);
                let result;
                Promise.resolve(op(store)).then(r => { result = r; }).catch(reject);
                tx.oncomplete = () => resolve(result);
                tx.onabort = tx.onerror = () => reject(tx.error || new Error('IndexedDB transaction failed.'));
            });
        },
        async saveHandle(handle) {
            await this._tx('readwrite', store => new Promise((res, rej) => {
                const r = store.put(handle, IDB_HANDLE_KEY);
                r.onsuccess = () => res();
                r.onerror = () => rej(r.error);
            }));
        },
        async loadHandle() {
            return this._tx('readonly', store => new Promise((res, rej) => {
                const r = store.get(IDB_HANDLE_KEY);
                r.onsuccess = () => res(r.result || null);
                r.onerror = () => rej(r.error);
            }));
        },
        async deleteHandle() {
            await this._tx('readwrite', store => new Promise((res, rej) => {
                const r = store.delete(IDB_HANDLE_KEY);
                r.onsuccess = () => res();
                r.onerror = () => rej(r.error);
            }));
        }
    };

    // ============================================================
    //  ExcelIO 模組（File System Access API + SheetJS）
    // ============================================================
    /** 工作表名稱常數（不分大小寫比對） */
    const SHEET_PLAIN = 'Plain';
    const SHEET_EXTERNAL = 'External';
    const SHEET_SCHEDULED = 'Scheduled';
    const SHEET_TEMPLATES = 'Templates'; // 舊版單一工作表（向下相容）
    const SHEET_TRAVEL = 'TravelApproval';

    function getPickerWindow() {
        if (typeof unsafeWindow !== 'undefined' && typeof unsafeWindow.showOpenFilePicker === 'function') {
            return unsafeWindow;
        }
        return window;
    }

    const ExcelIO = {
        isSupported() {
            const w = getPickerWindow();
            return typeof w !== 'undefined' && typeof w.showOpenFilePicker === 'function';
        },
        /**
         * 透過 showOpenFilePicker 讓使用者選擇本地 .xlsx，回傳 FileSystemFileHandle。
         * 需要在使用者手勢（click）內呼叫。
         */
        async pickFile() {
            if (!this.isSupported()) {
                throw new Error('File System Access API is not supported in this browser.');
            }
            const w = getPickerWindow();
            const [handle] = await w.showOpenFilePicker.call(w, {
                multiple: false,
                excludeAcceptAllOption: false,
                types: [{
                    description: 'Excel Workbook',
                    accept: {
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                    },
                }],
            });
            return handle;
        },
        /**
         * 確保 handle 仍有讀取權限；可選擇主動向使用者請求權限。
         * @param {FileSystemFileHandle} handle
         * @param {boolean} canPrompt  是否允許彈出權限請求（需在使用者手勢內）
         * @returns {Promise<'granted'|'denied'|'prompt'>}
         */
        async ensureReadPermission(handle, canPrompt) {
            if (!handle || typeof handle.queryPermission !== 'function') return 'granted';
            const opts = { mode: 'read' };
            const current = await handle.queryPermission(opts);
            if (current === 'granted') return 'granted';
            if (canPrompt && typeof handle.requestPermission === 'function') {
                const next = await handle.requestPermission(opts);
                return next;
            }
            return current;
        },
        /**
         * 讀取 handle 指向的檔案內容並解析成 workbook。
         * @param {FileSystemFileHandle} handle
         */
        async readWorkbook(handle) {
            if (typeof XLSX === 'undefined' || !XLSX || !XLSX.read) {
                throw new Error('SheetJS (XLSX) library not loaded.');
            }
            const file = await handle.getFile();
            const buf = await file.arrayBuffer();
            // SheetJS options：
            //   cellDates: false    — 我們自己處理日期格式，避免解讀為 Date 物件
            //   cellNF: false       — 不需要 number format
            //   cellText: false     — 用原始值即可
            return XLSX.read(buf, { type: 'array', cellDates: false, cellNF: false, cellText: false });
        }
    };

    // ============================================================
    //  ExcelParser 模組（Workbook → templates 資料結構）
    // ============================================================
    /**
     * Templates 欄位 → 模板物件的對映。Key = 標頭名稱（小寫、移除空白與底線後比對）。
     * Value = (tpl, rawValue) => 寫入 tpl
     */
    function normalizeHeaderName(name) {
        return String(name || '').trim().toLowerCase().replace(/[\s_]+/g, '');
    }

    function parseBool(v, fallback) {
        if (v == null || v === '') return fallback;
        if (typeof v === 'boolean') return v;
        const s = String(v).trim().toLowerCase();
        if (['true', 'yes', 'y', '1', 'on'].includes(s)) return true;
        if (['false', 'no', 'n', '0', 'off'].includes(s)) return false;
        return fallback;
    }

    /** 將多行字串解析為 key=value 物件（每行一組；忽略空行與 # 註解） */
    function parseKeyValueLines(text) {
        const out = {};
        if (!text) return out;
        String(text).split(/\r?\n/).forEach(line => {
            const trimmed = line.trim();
            if (!trimmed || trimmed.startsWith('#')) return;
            const eq = trimmed.indexOf('=');
            if (eq < 0) return;
            const k = trimmed.slice(0, eq).trim();
            const v = trimmed.slice(eq + 1).trim();
            if (k) out[k] = v;
        });
        return out;
    }

    /** 將多行字串解析為陣列（去除空行；可用 # 註解） */
    function parseLineList(text) {
        if (!text) return [];
        return String(text).split(/\r?\n/)
            .map(l => l.trim())
            .filter(l => l && !l.startsWith('#'));
    }

    /** 找到工作簿中名稱不分大小寫匹配的工作表 */
    function findSheet(workbook, name) {
        if (!workbook || !workbook.SheetNames) return null;
        const target = String(name).toLowerCase();
        const match = workbook.SheetNames.find(n => String(n).toLowerCase() === target);
        return match ? workbook.Sheets[match] : null;
    }

    /** 將工作表轉成 row 陣列（每 row 為以「正規化標頭」為 key 的物件） */
    function sheetToObjects(sheet) {
        if (!sheet) return [];
        // header:1 → 二維陣列；之後自己處理標頭以避免重複欄名問題
        const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            blankrows: false,
            defval: '',
            raw: true,
        });
        if (rows.length === 0) return [];
        const headers = rows[0].map(normalizeHeaderName);
        const out = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            const obj = {};
            let hasAny = false;
            for (let c = 0; c < headers.length; c++) {
                const key = headers[c];
                if (!key) continue;
                const val = row[c];
                obj[key] = val == null ? '' : val;
                if (val !== '' && val != null) hasAny = true;
            }
            if (hasAny) out.push(obj);
        }
        return out;
    }

    const ExcelParser = {
        /**
         * 將 workbook 轉成 { version: 1, templates: [...], travelApproval: {...} }
         * 並執行與原 YAML 相同的 schema 驗證。
         *
         * 支援兩種格式：
         *   新格式：workbook 含 Plain / External / Scheduled 分頁，每頁只含對應欄位
         *   舊格式：workbook 含 Templates 分頁，含 kind 欄及所有欄位（向下相容）
         *   若同時存在，優先使用新格式。
         *
         * @returns {{ ok: boolean, data?: Object, errors?: string[] }}
         */
        parse(workbook) {
            const errors = [];
            const plainSheet    = findSheet(workbook, SHEET_PLAIN);
            const externalSheet = findSheet(workbook, SHEET_EXTERNAL);
            const scheduledSheet = findSheet(workbook, SHEET_SCHEDULED);

            let templates;
            if (plainSheet || externalSheet || scheduledSheet) {
                // ── 新格式：每個 kind 各有獨立工作表 ──
                templates = [];
                if (plainSheet) {
                    sheetToObjects(plainSheet).forEach((row, idx) => {
                        templates.push(this._rowToTemplateForKind(row, 'plain', idx, errors, SHEET_PLAIN));
                    });
                }
                if (externalSheet) {
                    sheetToObjects(externalSheet).forEach((row, idx) => {
                        templates.push(this._rowToTemplateForKind(row, 'external', idx, errors, SHEET_EXTERNAL));
                    });
                }
                if (scheduledSheet) {
                    sheetToObjects(scheduledSheet).forEach((row, idx) => {
                        templates.push(this._rowToTemplateForKind(row, 'scheduled', idx, errors, SHEET_SCHEDULED));
                    });
                }
            } else {
                // ── 舊格式（向下相容）：單一 Templates 工作表 ──
                const templatesSheet = findSheet(workbook, SHEET_TEMPLATES);
                if (!templatesSheet) {
                    return { ok: false, errors: [
                        `Workbook is missing template sheets. Expected "${SHEET_PLAIN}", "${SHEET_EXTERNAL}", or "${SHEET_SCHEDULED}" (new format), or "${SHEET_TEMPLATES}" (legacy format).`
                    ] };
                }
                const rows = sheetToObjects(templatesSheet);
                templates = rows.map((row, idx) => this._rowToTemplate(row, idx, errors, SHEET_TEMPLATES));
            }
            const travelSheet = findSheet(workbook, SHEET_TRAVEL);
            const travelApproval = travelSheet ? this._parseTravelSheet(travelSheet, errors) : undefined;

            const data = {
                version: 1,
                templates,
                ...(travelApproval ? { travelApproval } : {}),
            };
            const { valid, errors: schemaErrors } = validateSchema(data);
            if (!valid) errors.push(...schemaErrors);
            if (errors.length > 0) return { ok: false, errors };
            return { ok: true, data };
        },

        _rowToTemplate(row, idx, errors, sheetName = SHEET_TEMPLATES) {
            const get = (k) => {
                const v = row[k];
                return v == null ? '' : (typeof v === 'string' ? v : String(v));
            };
            const id = get('id').trim();
            const kind = get('kind').trim().toLowerCase();
            const name = get('name').trim();
            const tpl = {
                id,
                kind,
                name,
                button: {
                    label: get('buttonlabel'),
                    icon: get('buttonicon'),
                    color: get('buttoncolor'),
                },
            };
            // Drop empty button sub-fields so validation messages stay focused on real issues
            ['label', 'icon', 'color'].forEach(k => {
                if (!tpl.button[k]) delete tpl.button[k];
            });
            if (Object.keys(tpl.button).length === 0) delete tpl.button;

            const orderRaw = row.order;
            if (orderRaw != null && orderRaw !== '') {
                const n = Number(orderRaw);
                if (Number.isInteger(n)) {
                    tpl.order = n;
                } else {
                    errors.push(`${sheetName} row ${idx + 2}: order must be an integer (not a decimal), got "${orderRaw}".`);
                }
            }

            const to = get('to');
            const cc = get('cc');
            const subject = get('subject');
            const body = get('body');
            if (kind === 'plain' || kind === 'scheduled') {
                tpl.to = to;
                tpl.cc = cc;
                tpl.subject = subject;
                tpl.body = body;
            }

            if (kind === 'external') {
                const ext = {
                    url: get('externalurl'),
                    params: parseKeyValueLines(get('externalparams')),
                };
                const openInNewTab = parseBool(row.externalopeninnewtab, undefined);
                if (openInNewTab !== undefined) ext.openInNewTab = openInNewTab;
                const addFRDate = parseBool(row.externaladdfrdate, undefined);
                if (addFRDate !== undefined) ext.addFRDate = addFRDate;
                if (Object.keys(ext.params).length === 0) delete ext.params;
                tpl.external = ext;

                const emailTo = get('emailto');
                const emailCc = get('emailcc');
                const emailSubject = get('emailsubject');
                const emailBody = get('emailbody');
                if (emailTo || emailCc || emailSubject || emailBody) {
                    tpl.email = {
                        to: emailTo,
                        cc: emailCc,
                        subject: emailSubject,
                        body: emailBody,
                    };
                }
            }

            if (kind === 'scheduled') {
                const dp = {
                    label: get('datepickerlabel'),
                    variable: get('datepickervariable'),
                    format: get('datepickerformat'),
                    min: get('datepickermin'),
                    max: get('datepickermax'),
                };
                const offsetRaw = row.datepickerdefaultoffsetdays;
                if (offsetRaw !== '' && offsetRaw != null) {
                    const n = parseInt(offsetRaw, 10);
                    if (!Number.isNaN(n)) dp.defaultOffsetDays = n;
                    else errors.push(`${sheetName} row ${idx + 2}: datePickerDefaultOffsetDays must be an integer.`);
                }
                Object.keys(dp).forEach(k => { if (dp[k] === '' || dp[k] == null) delete dp[k]; });
                tpl.datePicker = dp;
            }

            return tpl;
        },

        /**
         * 解析來自 kind 專屬工作表的一列資料（kind 由呼叫方傳入，列中不需要 kind 欄位）。
         * @param {Object} row         sheetToObjects() 回傳的列物件
         * @param {string} kind        'plain' | 'external' | 'scheduled'
         * @param {number} idx         零起始列索引（標頭不計）
         * @param {string[]} errors    累積錯誤陣列
         * @param {string} sheetName   工作表名稱（用於錯誤訊息）
         */
        _rowToTemplateForKind(row, kind, idx, errors, sheetName) {
            // 直接寫入 kind 欄位（sheetToObjects 每次回傳新物件，安全）
            row.kind = kind;
            return this._rowToTemplate(row, idx, errors, sheetName);
        },

        _parseTravelSheet(sheet, errors) {
            const rows = XLSX.utils.sheet_to_json(sheet, {
                header: 1,
                blankrows: false,
                defval: '',
                raw: true,
            });
            // Skip header row if first cell looks like a header (case-insensitive "field")
            let startIdx = 0;
            if (rows.length > 0) {
                const firstCell = String(rows[0][0] || '').trim().toLowerCase();
                if (firstCell === 'field' || firstCell === 'key' || firstCell === 'name') startIdx = 1;
            }
            const flat = {};
            for (let i = startIdx; i < rows.length; i++) {
                const row = rows[i];
                const key = String(row[0] || '').trim();
                if (!key || key.startsWith('#')) continue;
                flat[key] = row[1] == null ? '' : row[1];
            }

            const get = (path) => flat[path];
            const has = (path) => Object.prototype.hasOwnProperty.call(flat, path) && flat[path] !== '' && flat[path] != null;

            const ta = {};
            if (has('enabled')) ta.enabled = parseBool(get('enabled'), true);
            const button = {};
            if (has('button.label')) button.label = String(get('button.label'));
            if (has('button.icon')) button.icon = String(get('button.icon'));
            if (has('button.color')) button.color = String(get('button.color'));
            if (Object.keys(button).length) ta.button = button;

            const modal = {};
            if (has('modal.title')) modal.title = String(get('modal.title'));
            const sections = {};
            ['dates', 'locations', 'from', 'to', 'billing', 'amount', 'preview'].forEach(k => {
                const p = `modal.sections.${k}`;
                if (has(p)) sections[k] = String(get(p));
            });
            if (Object.keys(sections).length) modal.sections = sections;
            const placeholders = {};
            ['from', 'to', 'amount'].forEach(k => {
                const p = `modal.placeholders.${k}`;
                if (has(p)) placeholders[k] = String(get(p));
            });
            if (Object.keys(placeholders).length) modal.placeholders = placeholders;
            const labels = {};
            ['noDatesSelected', 'notFilled'].forEach(k => {
                const p = `modal.labels.${k}`;
                if (has(p)) labels[k] = String(get(p));
            });
            if (Object.keys(labels).length) modal.labels = labels;
            if (Object.keys(modal).length) ta.modal = modal;

            const options = {};
            if (has('options.locations')) {
                const list = parseLineList(get('options.locations'));
                if (list.length) options.locations = list;
            }
            if (has('options.billingTypes')) {
                const list = parseLineList(get('options.billingTypes'));
                if (list.length) options.billingTypes = list;
            }
            if (Object.keys(options).length) ta.options = options;

            const emailTemplates = {};
            if (has('emailTemplates.subject')) emailTemplates.subject = String(get('emailTemplates.subject'));
            if (has('emailTemplates.body')) emailTemplates.body = String(get('emailTemplates.body'));
            if (Object.keys(emailTemplates).length) ta.emailTemplates = emailTemplates;

            return Object.keys(ta).length ? ta : undefined;
        }
    };

    // ============================================================
    //  ExcelTemplateBuilder（用 DEFAULT_YAML 內容生成範例 .xlsx 供下載）
    // ============================================================
    const ExcelTemplateBuilder = {
        /** 將內建預設 YAML 轉成範例 workbook 並回傳 Blob */
        buildExampleBlob() {
            // 解析內建 YAML 作為示範資料
            const parsed = jsyaml.load(DEFAULT_YAML);
            const wb = XLSX.utils.book_new();

            // 依 kind 分類模板
            const allTemplates = parsed.templates || [];
            const byKind = { plain: [], external: [], scheduled: [] };
            allTemplates.forEach(tpl => {
                if (byKind[tpl.kind]) byKind[tpl.kind].push(tpl);
            });

            // ── Plain sheet ──
            const plainHeaders = [
                'id', 'name', 'order',
                'buttonLabel', 'buttonIcon', 'buttonColor',
                'to', 'cc', 'subject', 'body',
            ];
            const plainAoa = [plainHeaders];
            byKind.plain.forEach(tpl => {
                const btn = tpl.button || {};
                plainAoa.push([
                    tpl.id || '',
                    tpl.name || '',
                    tpl.order == null ? '' : tpl.order,
                    btn.label || '',
                    btn.icon || '',
                    btn.color || '',
                    tpl.to || '',
                    tpl.cc || '',
                    tpl.subject || '',
                    tpl.body || '',
                ]);
            });
            const plainWs = XLSX.utils.aoa_to_sheet(plainAoa);
            plainWs['!cols'] = plainHeaders.map(h => {
                if (['body', 'subject'].includes(h)) return { wch: 40 };
                return { wch: 16 };
            });
            XLSX.utils.book_append_sheet(wb, plainWs, SHEET_PLAIN);

            // ── External sheet ──
            const externalHeaders = [
                'id', 'name', 'order',
                'buttonLabel', 'buttonIcon', 'buttonColor',
                'externalUrl', 'externalParams', 'externalOpenInNewTab', 'externalAddFRDate',
                'emailTo', 'emailCc', 'emailSubject', 'emailBody',
            ];
            const externalAoa = [externalHeaders];
            byKind.external.forEach(tpl => {
                const btn = tpl.button || {};
                const ext = tpl.external || {};
                const email = tpl.email || {};
                const params = ext.params
                    ? Object.entries(ext.params).map(([k, v]) => `${k}=${v}`).join('\n')
                    : '';
                externalAoa.push([
                    tpl.id || '',
                    tpl.name || '',
                    tpl.order == null ? '' : tpl.order,
                    btn.label || '',
                    btn.icon || '',
                    btn.color || '',
                    ext.url || '',
                    params,
                    ext.openInNewTab == null ? '' : (ext.openInNewTab ? 'TRUE' : 'FALSE'),
                    ext.addFRDate == null ? '' : (ext.addFRDate ? 'TRUE' : 'FALSE'),
                    email.to || '',
                    email.cc || '',
                    email.subject || '',
                    email.body || '',
                ]);
            });
            const externalWs = XLSX.utils.aoa_to_sheet(externalAoa);
            externalWs['!cols'] = externalHeaders.map(h => {
                if (['externalUrl', 'externalParams', 'emailBody', 'emailSubject'].includes(h)) return { wch: 40 };
                return { wch: 16 };
            });
            XLSX.utils.book_append_sheet(wb, externalWs, SHEET_EXTERNAL);

            // ── Scheduled sheet ──
            const scheduledHeaders = [
                'id', 'name', 'order',
                'buttonLabel', 'buttonIcon', 'buttonColor',
                'to', 'cc', 'subject', 'body',
                'datePickerLabel', 'datePickerVariable', 'datePickerFormat',
                'datePickerDefaultOffsetDays', 'datePickerMin', 'datePickerMax',
            ];
            const scheduledAoa = [scheduledHeaders];
            byKind.scheduled.forEach(tpl => {
                const btn = tpl.button || {};
                const dp = tpl.datePicker || {};
                scheduledAoa.push([
                    tpl.id || '',
                    tpl.name || '',
                    tpl.order == null ? '' : tpl.order,
                    btn.label || '',
                    btn.icon || '',
                    btn.color || '',
                    tpl.to || '',
                    tpl.cc || '',
                    tpl.subject || '',
                    tpl.body || '',
                    dp.label || '',
                    dp.variable || '',
                    dp.format || '',
                    dp.defaultOffsetDays == null ? '' : dp.defaultOffsetDays,
                    dp.min || '',
                    dp.max || '',
                ]);
            });
            const scheduledWs = XLSX.utils.aoa_to_sheet(scheduledAoa);
            scheduledWs['!cols'] = scheduledHeaders.map(h => {
                if (['body', 'subject'].includes(h)) return { wch: 40 };
                if (['datePickerLabel'].includes(h)) return { wch: 26 };
                return { wch: 16 };
            });
            XLSX.utils.book_append_sheet(wb, scheduledWs, SHEET_SCHEDULED);

            // ── TravelApproval sheet ──
            const ta = parsed.travelApproval || {};
            const taRows = [['Field', 'Value']];
            const push = (k, v) => { if (v != null && v !== '') taRows.push([k, v]); };
            const boolStr = (v) => (v ? 'TRUE' : 'FALSE');
            if (ta.enabled != null) push('enabled', boolStr(ta.enabled));
            const btn = ta.button || {};
            push('button.label', btn.label);
            push('button.icon', btn.icon);
            push('button.color', btn.color);
            const modal = ta.modal || {};
            push('modal.title', modal.title);
            ['dates', 'locations', 'from', 'to', 'billing', 'amount', 'preview'].forEach(k => {
                push(`modal.sections.${k}`, modal.sections && modal.sections[k]);
            });
            ['from', 'to', 'amount'].forEach(k => {
                push(`modal.placeholders.${k}`, modal.placeholders && modal.placeholders[k]);
            });
            ['noDatesSelected', 'notFilled'].forEach(k => {
                push(`modal.labels.${k}`, modal.labels && modal.labels[k]);
            });
            const opts = ta.options || {};
            if (Array.isArray(opts.locations)) push('options.locations', opts.locations.join('\n'));
            if (Array.isArray(opts.billingTypes)) push('options.billingTypes', opts.billingTypes.join('\n'));
            const et = ta.emailTemplates || {};
            push('emailTemplates.subject', et.subject);
            push('emailTemplates.body', et.body);

            const taWs = XLSX.utils.aoa_to_sheet(taRows);
            taWs['!cols'] = [{ wch: 32 }, { wch: 60 }];
            XLSX.utils.book_append_sheet(wb, taWs, SHEET_TRAVEL);

            const arr = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
            return new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        },

        download() {
            const blob = this.buildExampleBlob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = TEMPLATE_FILE_NAME;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                URL.revokeObjectURL(url);
                a.remove();
            }, 1000);
        }
    };


    // ============================================================
    //  日期格式化 helper
    // ============================================================
    /**
     * 將 Date 物件格式化為指定格式。
     * @param {Date} date
     * @param {string} format  ISO | ROC | yyyy-MM-dd | yyyy/MM/dd
     * @returns {string}
     */
    function formatDate(date, format) {
        const y = date.getFullYear();
        const m = date.getMonth() + 1;
        const d = date.getDate();
        const mm = String(m).padStart(2, '0');
        const dd = String(d).padStart(2, '0');
        switch (format) {
            case 'ROC': {
                const rocYear = y - 1911;
                return `中華民國${rocYear}年${m}月${d}日`;
            }
            case 'yyyy/MM/dd':
                return `${y}/${mm}/${dd}`;
            case 'ISO':
            case 'yyyy-MM-dd':
            default:
                return `${y}-${mm}-${dd}`;
        }
    }

    // ============================================================
    //  TemplateEngine 模組
    // ============================================================
    const TemplateEngine = {
        /**
         * 將 template 字串中的 {{var}} 或 #var# 以 vars 物件替換。
         * 未匹配變數保留原樣。
         * @param {string} str
         * @param {Object} vars
         * @returns {string}
         */
        render(str, vars) {
            if (!str) return str || '';
            // 支援 {{var}}、{{ var }}（含空白）與舊式 #var#
            return str
                .replace(/\{\{\s*([\w]+)\s*\}\}/g, (_, key) =>
                    key in vars ? vars[key] : `{{${key}}}`)
                .replace(/#([\w]+)#/g, (_, key) =>
                    key in vars ? vars[key] : `#${key}#`);
        },

        /**
         * 以頁面資料與額外變數渲染整個 template 物件的文字欄位。
         * @param {Object} tpl  模板物件
         * @param {Object} extraVars  額外變數（例如 datePicker 選出的日期）
         * @returns {Object}  包含 to, cc, subject, body 的物件
         */
        renderAll(tpl, extraVars = {}) {
            const vars = Object.assign({}, PageData.gather(), extraVars);
            return {
                to:      this.render(tpl.to      || '', vars),
                cc:      this.render(tpl.cc      || '', vars),
                subject: this.render(tpl.subject || '', vars),
                body:    this.render(tpl.body    || '', vars),
            };
        }
    };

    const DEFAULT_TRAVEL_APPROVAL_CONFIG = Object.freeze({
        enabled: true,
        button: Object.freeze({
            label: 'Travel Approval',
            icon: '✈️',
            color: '#4A6F8A',
        }),
        modal: Object.freeze({
            title: '✈️ Travel Approval',
            sections: Object.freeze({
                dates: '📅 Travel Dates (Multiple Selection)',
                locations: '📍 Travel Locations',
                from: 'Departure Location:',
                to: 'Destination Location:',
                billing: '📊 Billing Type',
                amount: '💰 Estimated Travel Cost',
                preview: '📧 Email Preview',
            }),
            placeholders: Object.freeze({
                from: 'Enter departure location',
                to: 'Enter destination location',
                amount: 'Enter estimated amount (e.g. 2000)',
            }),
            labels: Object.freeze({
                noDatesSelected: '(No dates selected)',
                notFilled: '(not filled)',
            }),
        }),
        options: Object.freeze({
            locations: Object.freeze(['ULS Company site', '台北', '關渡賓士大樓', '群通大樓', '新北', '桃園', '新竹', '台中', '台南', '高雄', 'Customer Site']),
            billingTypes: Object.freeze(['Non-Billable', 'Billable Invoiceable', 'Billable Non-Invoiceable']),
        }),
        emailTemplates: Object.freeze({
            subject: 'Travel Approval - {{pjNum}} {{datesStr}}',
            body: [
                'Dear Manager,',
                '',
                'I would like to request travel approval with the following details:',
                '',
                'Project: {{pjNum}}{{pjNamePart}}',
                'Description: {{pjScope}}',
                '',
                'Travel Dates: {{datesStr}}',
                'Departure Location: {{fromLocation}}',
                'Destination Location: {{toLocation}}',
                '',
                'Billing Type: {{billingType}}',
                'Estimated Travel Cost: {{amount}}',
                '',
                'Please approve this request.',
                '',
                'Thank you.',
                '{{projectHandlerEmail}}',
            ].join('\n'),
        }),
    });

    function normalizeTravelApprovalConfig(raw) {
        const cfg = raw && typeof raw === 'object' ? raw : {};
        const button = cfg.button && typeof cfg.button === 'object' ? cfg.button : {};
        const modal = cfg.modal && typeof cfg.modal === 'object' ? cfg.modal : {};
        const modalSections = modal.sections && typeof modal.sections === 'object' ? modal.sections : {};
        const modalPlaceholders = modal.placeholders && typeof modal.placeholders === 'object' ? modal.placeholders : {};
        const modalLabels = modal.labels && typeof modal.labels === 'object' ? modal.labels : {};
        const options = cfg.options && typeof cfg.options === 'object' ? cfg.options : {};
        const emailTemplates = cfg.emailTemplates && typeof cfg.emailTemplates === 'object' ? cfg.emailTemplates : {};
        const locations = Array.isArray(options.locations) && options.locations.length > 0
            ? options.locations.map(v => String(v))
            : DEFAULT_TRAVEL_APPROVAL_CONFIG.options.locations.slice();
        const billingTypes = Array.isArray(options.billingTypes) && options.billingTypes.length > 0
            ? options.billingTypes.map(v => String(v))
            : DEFAULT_TRAVEL_APPROVAL_CONFIG.options.billingTypes.slice();

        return {
            enabled: cfg.enabled !== false,
            button: {
                label: typeof button.label === 'string' && button.label.trim() ? button.label : DEFAULT_TRAVEL_APPROVAL_CONFIG.button.label,
                icon: typeof button.icon === 'string' && button.icon.trim() ? button.icon : DEFAULT_TRAVEL_APPROVAL_CONFIG.button.icon,
                color: typeof button.color === 'string' && button.color.trim() ? button.color : DEFAULT_TRAVEL_APPROVAL_CONFIG.button.color,
            },
            modal: {
                title: typeof modal.title === 'string' && modal.title.trim() ? modal.title : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.title,
                sections: {
                    dates: typeof modalSections.dates === 'string' && modalSections.dates.trim() ? modalSections.dates : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.sections.dates,
                    locations: typeof modalSections.locations === 'string' && modalSections.locations.trim() ? modalSections.locations : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.sections.locations,
                    from: typeof modalSections.from === 'string' && modalSections.from.trim() ? modalSections.from : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.sections.from,
                    to: typeof modalSections.to === 'string' && modalSections.to.trim() ? modalSections.to : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.sections.to,
                    billing: typeof modalSections.billing === 'string' && modalSections.billing.trim() ? modalSections.billing : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.sections.billing,
                    amount: typeof modalSections.amount === 'string' && modalSections.amount.trim() ? modalSections.amount : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.sections.amount,
                    preview: typeof modalSections.preview === 'string' && modalSections.preview.trim() ? modalSections.preview : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.sections.preview,
                },
                placeholders: {
                    from: typeof modalPlaceholders.from === 'string' && modalPlaceholders.from.trim() ? modalPlaceholders.from : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.placeholders.from,
                    to: typeof modalPlaceholders.to === 'string' && modalPlaceholders.to.trim() ? modalPlaceholders.to : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.placeholders.to,
                    amount: typeof modalPlaceholders.amount === 'string' && modalPlaceholders.amount.trim() ? modalPlaceholders.amount : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.placeholders.amount,
                },
                labels: {
                    noDatesSelected: typeof modalLabels.noDatesSelected === 'string' && modalLabels.noDatesSelected.trim() ? modalLabels.noDatesSelected : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.labels.noDatesSelected,
                    notFilled: typeof modalLabels.notFilled === 'string' && modalLabels.notFilled.trim() ? modalLabels.notFilled : DEFAULT_TRAVEL_APPROVAL_CONFIG.modal.labels.notFilled,
                },
            },
            options: {
                locations,
                billingTypes,
            },
            emailTemplates: {
                subject: typeof emailTemplates.subject === 'string' && emailTemplates.subject.trim()
                    ? emailTemplates.subject
                    : DEFAULT_TRAVEL_APPROVAL_CONFIG.emailTemplates.subject,
                body: typeof emailTemplates.body === 'string' && emailTemplates.body.trim()
                    ? emailTemplates.body
                    : DEFAULT_TRAVEL_APPROVAL_CONFIG.emailTemplates.body,
            },
        };
    }

    // ============================================================
    //  PageData 模組（從頁面 DOM 抽取變數）
    // ============================================================
    const PageData = {
        /**
         * 從頁面 DOM 抽取變數，回傳變數 key/value 物件。
         * 使用 XPath 抽取 UL Portal 頁面的欄位資料。
         */
        gather() {
            // ── XPath helper：依 display-label-row 的文字尋找下方的 display-field-row ──
            function extractFieldByLabel(labelText) {
                const xpath = `//div[@class='display-label-row' and normalize-space(.)='${labelText}']/following-sibling::div[@class='display-field-row'][1]`;
                const el = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                return el ? el.textContent.trim() : '';
            }

            // ── XPath helper：Project Scope（含備用路徑）──
            function extractProjectScope() {
                let scope = extractFieldByLabel('Project Scope');
                if (!scope) {
                    const xpath = "//div[@class='div-product-attribute']//div[@class='display-label-row' and normalize-space(.)='Project Scope']/following-sibling::div[@class='project-scope-display'][1]";
                    const el = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (el) scope = el.textContent.trim();
                }
                return scope || '';
            }

            // ── XPath helper：Project Name（帶 customer-flag 與 ellipsis-ctrl class）──
            function extractProjectName() {
                const xpath = "//div[contains(@class, 'display-label-row') and contains(@class, 'customer-flag') and normalize-space(.)='Project Name']/following-sibling::div[contains(@class, 'display-field-row') and contains(@class, 'ellipsis-ctrl')][1]";
                const el = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                return el ? el.textContent.trim() : '';
            }

            // ── XPath helper：Order Number（從 dt/dd 結構抽取）──
            function extractOdrNum() {
                const xpath = "//dt[normalize-space(.)='Order Number:']/following-sibling::dd[1]//span";
                const el = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                return el ? el.textContent.trim() : '';
            }

            // ── XPath helper：Project Handler email（title 屬性優先）──
            function extractProjectHandlerEmail() {
                const xpath = "//dt[normalize-space(.)='Project Handler:']/following-sibling::dd[1]";
                const el = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                return el ? (el.getAttribute('title') || el.textContent.trim()) : '';
            }

            // ── XPath helper：Client Name 與 Client Email ──
            function extractClientInfo() {
                const DEFAULT_NAME  = 'N/A';
                const DEFAULT_EMAIL = 'N/A';
                const CONTACT_LABEL = 'Customer Company Contact';
                let clientName  = '';
                let clientEmail = '';

                const debug = (...args) => console.log('[FEG][extractClientInfo]', ...args);
                const emailRegex = /\b[A-Z0-9](?:[A-Z0-9._%+-]{0,62}[A-Z0-9])?@[A-Z0-9.-]+\.[A-Z]{2,}\b/i;
                const tryMatchEmail = (txt) => (txt || '').match(emailRegex)?.[0] || '';
                const escapeRegExp = (str) => str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

                const blockXpaths = [
                    `//div[@class='div-product-attribute' and .//div[@class='display-label-row' and normalize-space(.)='${CONTACT_LABEL}']]`,
                    `//div[contains(concat(' ', normalize-space(@class), ' '), ' div-product-attribute ') and .//div[contains(concat(' ', normalize-space(@class), ' '), ' display-label-row ') and normalize-space(.)='${CONTACT_LABEL}']]`,
                    `//div[contains(concat(' ', normalize-space(@class), ' '), ' div-product-attribute ') and .//div[contains(concat(' ', normalize-space(@class), ' '), ' display-label-row ') and contains(normalize-space(.), '${CONTACT_LABEL}')]]`,
                ];

                let block = null;
                for (const xpath of blockXpaths) {
                    block = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (block) {
                        debug('Matched block with XPath:', xpath);
                        break;
                    }
                }

                if (!block) {
                    const candidates = document.querySelectorAll('.div-product-attribute');
                    const contactLabelRegex = new RegExp(escapeRegExp(CONTACT_LABEL).replace(/\s+/g, '\\s+'), 'i');
                    block = Array.from(candidates).find((el) => contactLabelRegex.test(el.textContent || '')) || null;
                    if (block) debug('Matched block with querySelectorAll fallback.');
                }

                if (block) {
                    const allFields = block.querySelectorAll('.display-field-row');
                    const rows = Array.from(allFields).map(el => (el.textContent || '').trim()).filter(Boolean);
                    debug('Candidate rows:', rows);

                    for (const row of rows) {
                        const matchedEmail = tryMatchEmail(row);
                        if (matchedEmail) {
                            clientEmail = matchedEmail.trim();
                            break;
                        }
                    }

                    for (const row of rows) {
                        // 跳過含 email 的 row，避免把 "Email: xxx@xx.com" 當成客戶名稱。
                        if (emailRegex.test(row)) continue;
                        const contactPrefixRegex = new RegExp(`^${escapeRegExp(CONTACT_LABEL).replace(/\s+/g, '\\s+')}\\s*:?`, 'i');
                        const normalized = row.replace(contactPrefixRegex, '').replace(/^Name\s*:\s*/i, '').trim();
                        if (normalized) {
                            clientName = normalized;
                            break;
                        }
                    }

                    if (!clientEmail) {
                        clientEmail = tryMatchEmail(block.textContent || '');
                    }
                } else {
                    debug('Unable to locate Customer Company Contact block.');
                }

                if (!clientName || !clientEmail) {
                    debug('Extraction incomplete.', { clientName: clientName || DEFAULT_NAME, clientEmail: clientEmail || DEFAULT_EMAIL });
                } else {
                    debug('Extraction success.', { clientName, clientEmail });
                }

                return {
                    clientName: clientName || DEFAULT_NAME,
                    clientEmail: clientEmail || DEFAULT_EMAIL,
                };
            }

            // ── XPath helper：Project Hold Reason ──
            function extractProjectHoldReason() {
                return extractFieldByLabel('Project Hold Reason') || '';
            }

            // ── helper：將 MM/DD/YYYY 日期字串格式化為中華民國年月 ──
            function formatToROCYearMonth(dateStr) {
                if (!dateStr) return '';
                // 支援 MM/DD/YYYY 格式
                const match = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
                if (!match) return dateStr;
                const year  = parseInt(match[3], 10);
                const month = parseInt(match[1], 10);
                return `中華民國${year - 1911}年${month}月`;
            }

            // ── XPath helper：Date Booked 原始值（西元）──
            function extractDateBooked() {
                const xpath = "//div[@class='div-product-attribute']//div[@class='display-label-row' and normalize-space(.)='Date Booked']/following-sibling::div[@class='display-field-row'][1]";
                const el = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (!el) return '';
                return el.textContent.trim();
            }

            // ── XPath helper：Customer Company Address（Street, City, Country）──
            function extractCustomerAddress() {
                const xpath = "//div[@class='div-product-attribute' and .//div[contains(@class, 'display-label-row') and contains(@class, 'row-border-bottom') and normalize-space(.)='Customer Company Address']]";
                const block = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;

                let street = '';
                let city = '';
                let country = '';

                if (block) {
                    const fields = block.querySelectorAll('.display-field-row');
                    if (fields[0]) street = fields[0].textContent.trim();
                    if (fields[1]) {
                        const cityCountry = Array.from(fields[1].querySelectorAll('div'))
                            .map((el) => el.textContent.trim())
                            .filter(Boolean);
                        city = cityCountry[0] || '';
                        country = cityCountry[1] || '';
                    }
                }

                const parts = [street, city, country].filter(Boolean);
                return parts.join(', ') || '';
            }

            // ── 抽取各欄位 ──
            const pjNum               = extractFieldByLabel('Oracle Project Number');
            const odrNum              = extractOdrNum();
            const pjScope             = extractProjectScope();
            const pjName              = extractProjectName();
            const { clientName, clientEmail } = extractClientInfo();
            const projectHandlerEmail = extractProjectHandlerEmail();
            const projectHoldReason   = extractProjectHoldReason();
            const dateBooked          = extractDateBooked();
            const rocDateBooked       = formatToROCYearMonth(dateBooked);
            const customerAddress     = extractCustomerAddress();
            // ProjectID 用 /Project/Index/ 連結（非 /Project/Details/）
            const projectAnchorHref   = document.querySelector('a[href*="/Project/Index/"]')?.href
                                        || window.location.href;

            console.log('[FEG] Date/Address extraction:', { dateBooked, rocDateBooked, customerAddress });

            const vars = {
                pjNum,
                odrNum,
                pjScope,
                pjName,
                clientName,
                clientEmail,
                projectHandlerEmail,
                projectHoldReason,
                dateBooked,
                rocDateBooked,
                customerAddress,
                projectAnchorHref,
            };
            console.log('[FEG] gathered page data:', vars);
            return vars;
        }
    };

    // ============================================================
    //  YAML 驗證
    // ============================================================
    /**
     * 驗證解析後的 YAML 物件是否符合 schema。
     * @param {any} parsed
     * @returns {{ valid: boolean, errors: string[] }}
     */
    function validateSchema(parsed) {
        const errors = [];
        if (!parsed || typeof parsed !== 'object') {
            return { valid: false, errors: ['Root level must be an object.'] };
        }
        if (parsed.version !== 1) {
            errors.push(`version must be 1, got ${JSON.stringify(parsed.version)}.`);
        }
        if (!Array.isArray(parsed.templates)) {
            errors.push('templates must be an array.');
            return { valid: errors.length === 0, errors };
        }
        const seenIds = new Set();
        parsed.templates.forEach((tpl, idx) => {
            const prefix = `templates[${idx}]`;
            if (!tpl.id)   errors.push(`${prefix}: missing id.`);
            if (!tpl.name) errors.push(`${prefix}: missing name.`);
            if (!tpl.kind) errors.push(`${prefix}: missing kind.`);
            if (!['plain', 'external', 'scheduled'].includes(tpl.kind)) {
                errors.push(`${prefix}: kind must be plain / external / scheduled, got "${tpl.kind}".`);
            }
            if (tpl.id) {
                if (seenIds.has(tpl.id)) {
                    errors.push(`${prefix}: duplicate id "${tpl.id}".`);
                }
                seenIds.add(tpl.id);
            }
            if (tpl.order != null && !Number.isInteger(tpl.order)) {
                errors.push(`${prefix}: order must be an integer, got "${tpl.order}".`);
            }
            if (tpl.kind === 'plain') {
                if (!tpl.subject) errors.push(`${prefix} (plain): missing subject.`);
                if (!tpl.body)    errors.push(`${prefix} (plain): missing body.`);
            }
            if (tpl.kind === 'external') {
                if (!tpl.external || !tpl.external.url) {
                    errors.push(`${prefix} (external): missing external.url.`);
                }
            }
            if (tpl.kind === 'scheduled') {
                if (!tpl.datePicker || !tpl.datePicker.variable) {
                    errors.push(`${prefix} (scheduled): missing datePicker.variable.`);
                }
                if (!tpl.subject) errors.push(`${prefix} (scheduled): missing subject.`);
                if (!tpl.body)    errors.push(`${prefix} (scheduled): missing body.`);
            }
        });
        if (parsed.travelApproval != null) {
            const ta = parsed.travelApproval;
            if (!ta || typeof ta !== 'object' || Array.isArray(ta)) {
                errors.push('travelApproval must be an object.');
            } else {
                if (ta.enabled != null && typeof ta.enabled !== 'boolean') {
                    errors.push('travelApproval.enabled must be a boolean.');
                }
                if (ta.emailTemplates != null) {
                    const et = ta.emailTemplates;
                    if (!et || typeof et !== 'object' || Array.isArray(et)) {
                        errors.push('travelApproval.emailTemplates must be an object.');
                    } else {
                        if (et.subject != null && typeof et.subject !== 'string') {
                            errors.push('travelApproval.emailTemplates.subject must be a string.');
                        }
                        if (et.body != null && typeof et.body !== 'string') {
                            errors.push('travelApproval.emailTemplates.body must be a string.');
                        }
                    }
                }
            }
        }
        return { valid: errors.length === 0, errors };
    }

    // ============================================================
    //  TemplateLoader 模組（讀取與快取模板資料）
    // ============================================================
    /** 取出已驗證的 default 模式內建資料（解析失敗則返回空模板集） */
    function loadDefaultModeData() {
        try {
            const parsed = jsyaml.load(DEFAULT_MODE_YAML);
            const { valid, errors } = validateSchema(parsed);
            if (!valid) {
                console.warn('[FEGX] 內建 default 模板 YAML 有誤：', errors);
                return { version: 1, templates: [] };
            }
            return parsed;
        } catch (e) {
            console.error('[FEGX] 解析內建 default YAML 失敗：', e);
            return { version: 1, templates: [] };
        }
    }

    /** 空白資料結構，作為 customize 尚未綁定 / 權限未授予時的安全回傳值 */
    const EMPTY_DATA = Object.freeze({ version: 1, templates: [] });

    const TemplateLoader = {
        // 上次解析失敗的錯誤訊息（供 UI 顯示）
        lastError: null,

        getTemplateMode() {
            const mode = Storage.loadTemplateMode();
            if (mode === TEMPLATE_MODE_DEFAULT || mode === TEMPLATE_MODE_CUSTOMIZE) return mode;
            Storage.saveTemplateMode(TEMPLATE_MODE_DEFAULT);
            return TEMPLATE_MODE_DEFAULT;
        },

        /**
         * 載入目前模式對應的模板資料。
         * - default     → 內建 YAML 同步解析
         * - customize   → 從 IndexedDB 取出 handle，讀取 .xlsx，解析
         *                 若沒有 handle / 沒有讀取權限 → 回傳空資料 + lastError
         * @param {Object} [opts]
         * @param {boolean} [opts.canPrompt]  允許向使用者請求權限（需在使用者手勢內）
         * @returns {Promise<Object>} 模板資料
         */
        async loadTemplates(opts) {
            this.lastError = null;
            const mode = this.getTemplateMode();
            if (mode === TEMPLATE_MODE_DEFAULT) {
                return loadDefaultModeData();
            }
            // customize mode → Excel
            if (!ExcelIO.isSupported()) {
                this.lastError = 'File System Access API is not supported in this browser.';
                return EMPTY_DATA;
            }
            let handle;
            try {
                handle = await ExcelStore.loadHandle();
            } catch (e) {
                this.lastError = `Failed to read IndexedDB: ${e.message}`;
                return EMPTY_DATA;
            }
            if (!handle) {
                this.lastError = 'No Excel file selected. Click "📂 Pick Excel" to choose a local .xlsx.';
                return EMPTY_DATA;
            }
            const canPrompt = !!(opts && opts.canPrompt);
            let perm;
            try {
                perm = await ExcelIO.ensureReadPermission(handle, canPrompt);
            } catch (e) {
                this.lastError = `Permission check failed: ${e.message}`;
                return EMPTY_DATA;
            }
            if (perm !== 'granted') {
                this.lastError = 'Read permission not granted. Click "🔄 Reload" to re-grant access to the Excel file.';
                return EMPTY_DATA;
            }
            let workbook;
            try {
                workbook = await ExcelIO.readWorkbook(handle);
            } catch (e) {
                this.lastError = `Failed to read Excel file: ${e.message}`;
                return EMPTY_DATA;
            }
            const result = ExcelParser.parse(workbook);
            if (!result.ok) {
                this.lastError = result.errors.join('\n');
                console.warn('[FEGX] Excel parse errors:', result.errors);
                return EMPTY_DATA;
            }
            return result.data;
        },

        /** 取得目前綁定的檔案資訊（顯示用），不會請求權限 */
        async getBoundFileInfo() {
            try {
                const handle = await ExcelStore.loadHandle();
                if (!handle) return null;
                let perm = 'unknown';
                try { perm = await ExcelIO.ensureReadPermission(handle, false); } catch (_) { /* ignore */ }
                return { name: handle.name || '(unknown)', permission: perm };
            } catch (e) {
                return null;
            }
        }
    };

    // ============================================================
    //  mailto 工具函式
    // ============================================================
    function openMailto(to, cc, subject, body) {
        const params = [];
        if (cc)      params.push(`cc=${encodeURIComponent(cc)}`);
        if (subject) params.push(`subject=${encodeURIComponent(subject)}`);
        if (body)    params.push(`body=${encodeURIComponent(body)}`);
        const query = params.join('&');
        window.location.href = `mailto:${encodeURIComponent(to)}${query ? '?' + query : ''}`;
    }

    // ============================================================
    //  DatePickerModal 模組
    // ============================================================
    const DatePickerModal = {
        /**
         * 顯示日期選擇 modal。
         * @param {Object} datePicker  YAML 中的 datePicker 設定
         * @param {Function} onConfirm  callback(dateStr: string, rawDate: Date)
         * @param {Function} onCancel   callback()
         */
        show(datePicker, onConfirm, onCancel) {
            // 計算預設日期
            const today = new Date();
            const offset = parseInt(datePicker.defaultOffsetDays, 10) || 0;
            const defaultDate = new Date(today.getTime() + offset * 86400000);
            const isoDefault = formatDate(defaultDate, 'yyyy-MM-dd');

            // 計算 min
            let minAttr = '';
            if (datePicker.min === 'today' || datePicker.min === 'Today') {
                minAttr = formatDate(today, 'yyyy-MM-dd');
            } else if (datePicker.min) {
                minAttr = datePicker.min;
            }
            const maxAttr = datePicker.max || '';

            // 建立 overlay
            const overlay = document.createElement('div');
            overlay.className = 'feg-modal-overlay';

            const dialog = document.createElement('div');
            dialog.className = 'feg-modal-dialog';

            const title = document.createElement('div');
            title.className = 'feg-modal-title';
            title.textContent = datePicker.label || 'Select Date';

            const input = document.createElement('input');
            input.type  = 'date';
            input.className = 'feg-modal-date-input';
            input.value = isoDefault;
            if (minAttr) input.min = minAttr;
            if (maxAttr) input.max = maxAttr;

            const btnRow = document.createElement('div');
            btnRow.className = 'feg-modal-btn-row';

            const btnConfirm = document.createElement('button');
            btnConfirm.className = 'feg-btn feg-btn-primary';
            btnConfirm.textContent = '✔ Confirm';

            const btnCancel = document.createElement('button');
            btnCancel.className = 'feg-btn feg-btn-secondary';
            btnCancel.textContent = '✕ Cancel';

            btnConfirm.addEventListener('click', () => {
                const val = input.value;
                if (!val) {
                    UI.toast('Please select a date.', 'warn');
                    return;
                }
                const parts = val.split('-');
                const rawDate = new Date(
                    parseInt(parts[0], 10),
                    parseInt(parts[1], 10) - 1,
                    parseInt(parts[2], 10)
                );
                const dateStr = formatDate(rawDate, datePicker.format || 'yyyy-MM-dd');
                overlay.remove();
                onConfirm(dateStr, rawDate);
            });

            btnCancel.addEventListener('click', () => {
                overlay.remove();
                if (onCancel) onCancel();
            });

            btnRow.append(btnCancel, btnConfirm);
            dialog.append(title, input, btnRow);
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            input.focus();
        }
    };

    // ============================================================
    //  TravelApprovalModal 模組
    // ============================================================
    const TravelApprovalModal = {
        show(rawConfig) {
            const config = normalizeTravelApprovalConfig(rawConfig);
            const LOCATIONS = config.options.locations;
            const BILLING_TYPES = config.options.billingTypes;

            const pageData  = PageData.gather();
            const today     = new Date();
            let calYear     = today.getFullYear();
            let calMonth    = today.getMonth();
            let selectedDates = [];          // 'YYYY-MM-DD' strings
            let billingType   = BILLING_TYPES[0];

            // ── Overlay ──────────────────────────────────────────
            const overlay = document.createElement('div');
            overlay.className = 'feg-modal-overlay';

            const dialog = document.createElement('div');
            dialog.className = 'feg-modal-dialog feg-ta-dialog';

            // ── Title ────────────────────────────────────────────
            const titleEl = document.createElement('div');
            titleEl.className = 'feg-modal-title';
            titleEl.textContent = config.modal.title;

            // ── Calendar Section ─────────────────────────────────
            const calSection = document.createElement('div');
            calSection.className = 'feg-ta-section';

            const calLabel = document.createElement('div');
            calLabel.className = 'feg-ta-label';
            calLabel.textContent = config.modal.sections.dates;

            const calWrap = document.createElement('div');
            calWrap.className = 'feg-ta-calendar-wrap';

            const calHeader = document.createElement('div');
            calHeader.className = 'feg-ta-cal-header';

            const btnPrev = document.createElement('button');
            btnPrev.type = 'button';
            btnPrev.className = 'feg-btn feg-btn-secondary feg-ta-cal-nav';
            btnPrev.textContent = '◀';

            const calMonthLabel = document.createElement('span');
            calMonthLabel.className = 'feg-ta-cal-month-label';

            const btnNext = document.createElement('button');
            btnNext.type = 'button';
            btnNext.className = 'feg-btn feg-btn-secondary feg-ta-cal-nav';
            btnNext.textContent = '▶';

            calHeader.append(btnPrev, calMonthLabel, btnNext);

            const calGrid = document.createElement('div');
            calGrid.className = 'feg-ta-cal-grid';

            calWrap.append(calHeader, calGrid);

            const selDatesWrap = document.createElement('div');
            selDatesWrap.className = 'feg-ta-selected-dates';

            calSection.append(calLabel, calWrap, selDatesWrap);

            // ── Location Section ─────────────────────────────────
            const locSection = document.createElement('div');
            locSection.className = 'feg-ta-section';

            const locLabel = document.createElement('div');
            locLabel.className = 'feg-ta-label';
            locLabel.textContent = config.modal.sections.locations;

            const locRow = document.createElement('div');
            locRow.className = 'feg-ta-loc-row';
            const modalListKey = ++travelLocListIdSeq;
            let locInputSeq = 0;

            function buildLocInput(labelText, options, defaultVal, placeholder) {
                const wrap = document.createElement('div');
                wrap.className = 'feg-ta-loc-field';
                const lbl = document.createElement('label');
                lbl.className = 'feg-ta-sublabel';
                lbl.textContent = labelText;
                const input = document.createElement('input');
                input.type = 'text';
                input.className = 'feg-ta-input';
                input.placeholder = placeholder || 'Type a location';
                if (defaultVal) input.value = defaultVal;
                const listId = `feg-ta-loc-list-${modalListKey}-${locInputSeq++}`;
                input.setAttribute('list', listId);
                const list = document.createElement('datalist');
                list.id = listId;
                options.forEach(o => {
                    const opt = document.createElement('option');
                    opt.value = o;
                    list.appendChild(opt);
                });
                wrap.append(lbl, input, list);
                return { wrap, input };
            }

            const { wrap: fromWrap, input: fromInput } = buildLocInput(
                config.modal.sections.from,
                LOCATIONS,
                LOCATIONS[0],
                config.modal.placeholders.from
            );
            const defaultDestination = pageData.customerAddress
                || LOCATIONS.find(loc => String(loc).trim().toLowerCase() === 'customer site')
                || LOCATIONS[0]
                || '';
            const { wrap: toWrap,   input: toInput   } = buildLocInput(
                config.modal.sections.to,
                LOCATIONS,
                defaultDestination,
                config.modal.placeholders.to
            );

            locRow.append(fromWrap, toWrap);
            locSection.append(locLabel, locRow);

            fromInput.addEventListener('input', updatePreview);
            toInput.addEventListener('input', updatePreview);

            // ── Billing Section ──────────────────────────────────
            const billSection = document.createElement('div');
            billSection.className = 'feg-ta-section';

            const billLabel = document.createElement('div');
            billLabel.className = 'feg-ta-label';
            billLabel.textContent = config.modal.sections.billing;

            const billGroup = document.createElement('div');
            billGroup.className = 'feg-ta-radio-group';

            BILLING_TYPES.forEach((type, idx) => {
                const lbl = document.createElement('label');
                lbl.className = 'feg-ta-radio-label';
                const radio = document.createElement('input');
                radio.type = 'radio';
                radio.name = 'feg-ta-billing';
                radio.value = type;
                if (idx === 0) radio.checked = true;
                radio.addEventListener('change', () => { billingType = type; updatePreview(); });
                lbl.append(radio, document.createTextNode(' ' + type));
                billGroup.appendChild(lbl);
            });

            billSection.append(billLabel, billGroup);

            // ── Amount Section ───────────────────────────────────
            const amtSection = document.createElement('div');
            amtSection.className = 'feg-ta-section';

            const amtLabel = document.createElement('div');
            amtLabel.className = 'feg-ta-label';
            amtLabel.textContent = config.modal.sections.amount;

            const amtInput = document.createElement('input');
            amtInput.type = 'text';
            amtInput.className = 'feg-ta-input';
            amtInput.placeholder = config.modal.placeholders.amount;
            amtInput.addEventListener('input', updatePreview);

            amtSection.append(amtLabel, amtInput);

            // ── Preview Section ──────────────────────────────────
            const prevSection = document.createElement('div');
            prevSection.className = 'feg-ta-section';

            const prevLabel = document.createElement('div');
            prevLabel.className = 'feg-ta-label';
            prevLabel.textContent = config.modal.sections.preview;

            const prevArea = document.createElement('textarea');
            prevArea.className = 'feg-ta-preview';
            prevArea.readOnly = true;

            prevSection.append(prevLabel, prevArea);

            // ── Helper functions ─────────────────────────────────
            function buildEmailContent() {
                const datesStr = selectedDates.length > 0
                    ? selectedDates.join(', ')
                    : config.modal.labels.noDatesSelected;
                const pjNum   = pageData.pjNum   || '';
                const pjName  = pageData.pjName  || '';
                const pjScope = pageData.pjScope || '';
                const handler = pageData.projectHandlerEmail || '';

                const NOT_FILLED = config.modal.labels.notFilled;
                const fromLocation = fromInput.value.trim() || NOT_FILLED;
                const toLocation = toInput.value.trim() || NOT_FILLED;
                const amount = amtInput.value.trim() || NOT_FILLED;
                const templateVars = {
                    pjNum,
                    pjName,
                    pjNamePart: pjName ? ` - ${pjName}` : '',
                    pjScope,
                    projectHandlerEmail: handler,
                    datesStr,
                    fromLocation,
                    toLocation,
                    billingType,
                    amount,
                };
                const subject = TemplateEngine.render(config.emailTemplates.subject, templateVars).trim();
                const body = TemplateEngine.render(config.emailTemplates.body, templateVars);

                return { subject, body };
            }

            function updatePreview() {
                const { subject, body } = buildEmailContent();
                prevArea.value = `Subject: ${subject}\n\n${body}`;
            }

            function renderCalendar() {
                const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
                calMonthLabel.textContent = `${MONTH_NAMES[calMonth]} ${calYear}`;

                calGrid.innerHTML = '';

                ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].forEach(d => {
                    const cell = document.createElement('div');
                    cell.className = 'feg-ta-cal-day-header';
                    cell.textContent = d;
                    calGrid.appendChild(cell);
                });

                const firstWeekday = new Date(calYear, calMonth, 1).getDay();
                const daysInMonth  = new Date(calYear, calMonth + 1, 0).getDate();

                for (let i = 0; i < firstWeekday; i++) {
                    calGrid.appendChild(document.createElement('div'));
                }

                for (let day = 1; day <= daysInMonth; day++) {
                    const ds = `${calYear}-${String(calMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                    const cell = document.createElement('div');
                    cell.className = 'feg-ta-cal-day';
                    cell.textContent = day;
                    if (selectedDates.includes(ds)) cell.classList.add('feg-ta-cal-day-selected');
                    const isToday = (calYear === today.getFullYear() && calMonth === today.getMonth() && day === today.getDate());
                    if (isToday) cell.classList.add('feg-ta-cal-day-today');
                    cell.addEventListener('click', () => {
                        const idx = selectedDates.indexOf(ds);
                        if (idx >= 0) selectedDates.splice(idx, 1);
                        else { selectedDates.push(ds); selectedDates.sort(); }
                        renderCalendar();
                        renderSelectedDates();
                        updatePreview();
                    });
                    calGrid.appendChild(cell);
                }
            }

            function renderSelectedDates() {
                selDatesWrap.innerHTML = '';
                if (selectedDates.length === 0) {
                    const hint = document.createElement('span');
                    hint.style.cssText = 'font-size:12px;color:#888;';
                    hint.textContent = config.modal.labels.noDatesSelected;
                    selDatesWrap.appendChild(hint);
                } else {
                    selectedDates.forEach(ds => {
                        const tag = document.createElement('span');
                        tag.className = 'feg-ta-date-tag';
                        tag.textContent = `${ds} ×`;
                        tag.addEventListener('click', () => {
                            selectedDates = selectedDates.filter(d => d !== ds);
                            renderCalendar();
                            renderSelectedDates();
                            updatePreview();
                        });
                        selDatesWrap.appendChild(tag);
                    });
                }
            }

            btnPrev.addEventListener('click', () => {
                calMonth--; if (calMonth < 0) { calMonth = 11; calYear--; }
                renderCalendar();
            });
            btnNext.addEventListener('click', () => {
                calMonth++; if (calMonth > 11) { calMonth = 0; calYear++; }
                renderCalendar();
            });

            // ── Button Row ───────────────────────────────────────
            const btnRow = document.createElement('div');
            btnRow.className = 'feg-modal-btn-row';

            const btnCancel = document.createElement('button');
            btnCancel.type = 'button';
            btnCancel.className = 'feg-btn feg-btn-secondary';
            btnCancel.textContent = '✕ Cancel';
            btnCancel.addEventListener('click', () => overlay.remove());

            const btnCopy = document.createElement('button');
            btnCopy.type = 'button';
            btnCopy.className = 'feg-btn feg-btn-secondary';
            btnCopy.textContent = '📋 Copy';
            btnCopy.addEventListener('click', () => {
                try {
                    const { subject, body } = buildEmailContent();
                    GM_setClipboard(`Subject: ${subject}\n\n${body}`, 'text');
                    UI.toast('✅ Copied to clipboard.', 'success');
                } catch (e) {
                    UI.toast(`Copy failed: ${e.message}`, 'error');
                }
            });

            const btnSend = document.createElement('button');
            btnSend.type = 'button';
            btnSend.className = 'feg-btn feg-btn-primary';
            btnSend.textContent = '✉️ Generate Email';
            btnSend.addEventListener('click', () => {
                if (selectedDates.length === 0) {
                    UI.toast('Please select at least one travel date.', 'warn');
                    return;
                }
                const { subject, body } = buildEmailContent();
                overlay.remove();
                openMailto('', '', subject, body);
            });

            btnRow.append(btnCancel, btnCopy, btnSend);

            // ── Assemble Dialog ──────────────────────────────────
            dialog.append(titleEl, calSection, locSection, billSection, amtSection, prevSection, btnRow);
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);

            // Initial render
            renderCalendar();
            renderSelectedDates();
            updatePreview();
        }
    };

    // ============================================================
    //  按鈕動作處理
    // ============================================================
    function handlePlain(tpl) {
        const rendered = TemplateEngine.renderAll(tpl);
        openMailto(rendered.to, rendered.cc, rendered.subject, rendered.body);
    }

    function handleExternal(tpl) {
        const vars = PageData.gather();
        const ext  = tpl.external;

        // 組 URL
        const url   = new URL(TemplateEngine.render(ext.url, vars));
        if (ext.params) {
            Object.entries(ext.params).forEach(([k, v]) => {
                url.searchParams.set(k, TemplateEngine.render(String(v), vars));
            });
        }
        if (ext.addFRDate) {
            const today = new Date();
            const frDate = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
            url.searchParams.set('FRDate', frDate);
        }

        if (ext.openInNewTab !== false) {
            window.open(url.toString(), '_blank');
        } else {
            window.location.href = url.toString();
        }

        // 若有 email 區塊，同時開 mailto
        if (tpl.email) {
            const emailRendered = {
                to:      TemplateEngine.render(tpl.email.to      || '', vars),
                cc:      TemplateEngine.render(tpl.email.cc      || '', vars),
                subject: TemplateEngine.render(tpl.email.subject || '', vars),
                body:    TemplateEngine.render(tpl.email.body    || '', vars),
            };
            // 小延遲，避免兩個視窗動作衝突
            setTimeout(() => {
                openMailto(emailRendered.to, emailRendered.cc, emailRendered.subject, emailRendered.body);
            }, 300);
        }
    }

    function handleScheduled(tpl) {
        DatePickerModal.show(
            tpl.datePicker,
            (dateStr) => {
                const extra = { [tpl.datePicker.variable]: dateStr };
                const rendered = TemplateEngine.renderAll(tpl, extra);
                openMailto(rendered.to, rendered.cc, rendered.subject, rendered.body);
            }
        );
    }

    // ============================================================
    //  UI 模組
    // ============================================================
    const UI = {
        _panel: null,
        _visible: true,
        _expanded: false,
        _hovering: false,
        _hoverCloseTimer: null,
        _templates: [],
        _templateMode: TEMPLATE_MODE_DEFAULT,
        _travelApprovalConfig: normalizeTravelApprovalConfig(null),

        /** 初始化整個 UI */
        init(data) {
            this._templates = data.templates || [];
            this._templateMode = TemplateLoader.getTemplateMode();
            this._travelApprovalConfig = normalizeTravelApprovalConfig(data.travelApproval);
            this._injectStyles();
            this._buildPanel();
        },

        /** 更新模板列表並重建面板 */
        reload(data) {
            this._templates = (data && data.templates) || [];
            this._templateMode = TemplateLoader.getTemplateMode();
            this._travelApprovalConfig = normalizeTravelApprovalConfig(data && data.travelApproval);
            const old = document.getElementById(PANEL_ID);
            if (old) old.remove();
            this._buildPanel();
            this.toast(`✅ ${this._templateMode === TEMPLATE_MODE_DEFAULT ? 'Default' : 'Cstm.'} templates loaded.`, 'success');
        },

        /** 以 console 輸出訊息（取代畫面 toast） */
        toast(msg, type = 'info') {
            const text = `[FEGX][${type}] ${msg}`;
            if (type === 'error') {
                console.error(text);
            } else if (type === 'warn') {
                console.warn(text);
            } else {
                console.log(text);
            }
        },

        togglePanelVisibility() {
            this._visible = !this._visible;
            this._applyPanelVisibility();
        },

        togglePanelExpanded() {
            this._expanded = !this._expanded;
            this._applyPanelExpandedState();
        },

        // ---- 內部方法 ----

        _injectStyles() {
            const style = document.createElement('style');
            style.textContent = `
/* ========= Flex Email Generator (feg-) ========= */
#feg-panel {
    --feg-panel-min-width: 220px;
    --feg-panel-max-width: min(320px, calc(100vw - 24px));
    --feg-panel-max-height: min(70vh, 520px);
    --feg-panel-transition: .16s ease;
    position: fixed;
    right: 12px;
    bottom: 12px;
    z-index: 999999;
    display: flex;
    flex-direction: column-reverse;
    align-items: flex-end;
    gap: 8px;
    max-width: calc(100vw - 24px);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    font-size: 13px;
    overflow: visible;
    user-select: none;
}
#feg-panel.feg-panel-hovering .feg-body,
#feg-panel:focus-within .feg-body,
#feg-panel.feg-panel-open .feg-body {
    opacity: 1;
    visibility: visible;
    pointer-events: auto;
    transform: translateY(0);
}
.feg-header {
    display: flex;
    align-items: center;
    gap: 5px;
    padding: 5px 10px;
    background: #fff;
    border: 1px solid #dde3ec;
    border-radius: 999px;
    box-shadow: 0 6px 28px rgba(0,0,0,.18);
    color: #3a4252;
    white-space: nowrap;
    cursor: pointer;
}
.feg-header-title {
    font-weight: 700;
    font-size: 12px;
    line-height: 1.2;
}
.feg-header-hint {
    font-size: 11px;
    color: #6b7280;
    line-height: 1.2;
}
.feg-header-btn {
    background: none;
    border: 1px solid #c9d1de;
    border-radius: 7px;
    padding: 3px 7px;
    font-size: 12px;
    cursor: pointer;
    color: #505a72;
    line-height: 1.1;
    transition: background .1s;
}
.feg-header-btn:hover { background: #e8ecf5; }
.feg-header-btn:disabled {
    background: #f3f4f6;
    border-color: #d1d5db;
    color: #9ca3af;
    cursor: not-allowed;
    opacity: 0.7;
    box-shadow: none;
}
.feg-header-btn:disabled:hover { background: #f3f4f6; }
.feg-mode-toggle.feg-mode-default {
    background: #d1d5db;
    border-color: #9ca3af;
    color: #374151;
}
.feg-body {
    display: flex;
    flex-direction: column;
    align-items: stretch;
    gap: 6px;
    min-width: var(--feg-panel-min-width);
    max-width: var(--feg-panel-max-width);
    max-height: var(--feg-panel-max-height);
    padding: 8px;
    background: #fff;
    border: 1px solid #dde3ec;
    border-radius: 12px;
    box-shadow: 0 6px 28px rgba(0,0,0,.18);
    overflow-x: hidden;
    overflow-y: auto;
    opacity: 0;
    visibility: hidden;
    pointer-events: none;
    transform: translateY(8px);
    transition: opacity var(--feg-panel-transition), transform var(--feg-panel-transition), visibility var(--feg-panel-transition);
    position: absolute;
    right: 0;
    bottom: calc(100% + 8px);
}
.feg-action-row {
    display: flex;
    gap: 6px;
    flex-wrap: wrap;
}
.feg-toolbar-sep {
    flex: 0 0 auto;
    height: 1px;
    background: #dde3ec;
}
.feg-empty-message {
    color: #6b7280;
    font-size: 12px;
    line-height: 1.3;
}
.feg-action-btn {
    flex: 1 1 0;
    min-width: 0;
}
.feg-tpl-btn {
    display: flex;
    align-items: center;
    gap: 7px;
    width: 100%;
    padding: 4px 9px;
    border-radius: 7px;
    border: none;
    cursor: pointer;
    font-size: 12px;
    font-weight: 600;
    color: #fff;
    text-align: left;
    transition: filter .15s, transform .1s;
    box-shadow: 0 2px 6px rgba(0,0,0,.12);
    line-height: 1.2;
}
.feg-tpl-btn:hover  { filter: brightness(1.10); transform: translateY(-1px); }
.feg-tpl-btn:active { filter: brightness(0.92); transform: translateY(0); }
.feg-tpl-btn-icon { font-size: 14px; line-height: 1; }
.feg-sidebar-separator {
    list-style: none;
    margin-top: 10px;
    padding-top: 10px;
    border-top: 1px solid #dde3ec;
}
.feg-sidebar-group-title {
    list-style: none;
    font-size: 12px;
    font-weight: 700;
    color: #505a72;
    margin: 0 0 6px;
}
.feg-sidebar-link {
    display: flex;
    align-items: center;
    gap: 8px;
    width: 100%;
    box-sizing: border-box;
    text-decoration: none;
    cursor: pointer;
}
.feg-sidebar-link:hover,
.feg-sidebar-link:focus {
    text-decoration: none;
}
.feg-sidebar-icon {
    font-size: 15px;
    line-height: 1;
}
.feg-sidebar-label {
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
/* Modal Overlay */
.feg-modal-overlay {
    position: fixed;
    inset: 0;
    z-index: 1000001;
    background: rgba(0,0,0,.45);
    display: flex;
    align-items: center;
    justify-content: center;
}
.feg-modal-dialog {
    background: #fff;
    border-radius: 12px;
    padding: 24px;
    min-width: 300px;
    max-width: 480px;
    box-shadow: 0 8px 32px rgba(0,0,0,.22);
    display: flex;
    flex-direction: column;
    gap: 14px;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}
.feg-modal-title {
    font-size: 15px;
    font-weight: 700;
    color: #1e2a3a;
}
.feg-modal-date-input {
    width: 100%;
    padding: 8px 10px;
    border: 1px solid #c9d1de;
    border-radius: 7px;
    font-size: 14px;
    box-sizing: border-box;
}
.feg-modal-btn-row {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
}
.feg-modal-textarea {
    width: 100%;
    height: 520px !important;
    padding: 8px 10px;
    border: 1px solid #c9d1de;
    border-radius: 7px;
    font-size: 12.5px;
    font-family: "Cascadia Code", "Fira Code", monospace;
    resize: vertical;
    box-sizing: border-box;
}
.feg-modal-error {
    color: #c0392b;
    font-size: 12px;
    background: #fdf0ef;
    border: 1px solid #f5c6c2;
    border-radius: 6px;
    padding: 6px 10px;
    white-space: pre-wrap;
    max-height: 120px;
    overflow-y: auto;
    display: none;
}
.feg-btn {
    padding: 7px 16px;
    border-radius: 7px;
    border: none;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: filter .12s;
}
.feg-btn:hover { filter: brightness(1.08); }
.feg-btn-primary   { background: #1a56db; color: #fff; }
.feg-btn-secondary { background: #e8ecf5; color: #3a4252; }
.feg-btn-danger    { background: #c0392b; color: #fff; }
/* Travel Approval Modal */
.feg-ta-dialog {
    min-width: 500px;
    max-width: 620px;
    max-height: 88vh;
    overflow-y: auto;
}
.feg-ta-section {
    display: flex;
    flex-direction: column;
    gap: 6px;
}
.feg-ta-label {
    font-size: 13px;
    font-weight: 700;
    color: #1e2a3a;
}
.feg-ta-sublabel {
    font-size: 12px;
    color: #505a72;
}
.feg-ta-select, .feg-ta-input {
    width: 100%;
    padding: 7px 10px;
    border: 1px solid #c9d1de;
    border-radius: 7px;
    font-size: 13px;
    box-sizing: border-box;
}
.feg-ta-radio-group {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}
.feg-ta-radio-label {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    font-size: 13px;
    cursor: pointer;
    padding: 5px 12px;
    border: 1px solid #c9d1de;
    border-radius: 7px;
    user-select: none;
}
.feg-ta-radio-label:has(input:checked) {
    background: #e8f0fe;
    border-color: #1a56db;
    color: #1a56db;
    font-weight: 600;
}
.feg-ta-loc-row {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
}
.feg-ta-loc-field {
    display: flex;
    flex-direction: column;
    gap: 4px;
}
.feg-ta-calendar-wrap {
    border: 1px solid #c9d1de;
    border-radius: 8px;
    overflow: hidden;
}
.feg-ta-cal-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 6px 10px;
    background: #f5f7fb;
    border-bottom: 1px solid #dde3ec;
}
.feg-ta-cal-month-label {
    font-weight: 700;
    font-size: 13px;
    color: #1e2a3a;
}
.feg-ta-cal-nav {
    padding: 2px 8px !important;
    font-size: 11px !important;
    min-width: 28px;
}
.feg-ta-cal-grid {
    display: grid;
    grid-template-columns: repeat(7, 1fr);
    gap: 2px;
    padding: 6px;
}
.feg-ta-cal-day-header {
    text-align: center;
    font-size: 11px;
    font-weight: 700;
    color: #505a72;
    padding: 3px 0;
}
.feg-ta-cal-day {
    text-align: center;
    padding: 5px 2px;
    border-radius: 5px;
    font-size: 12px;
    cursor: pointer;
    transition: background .1s;
}
.feg-ta-cal-day:hover { background: #e8ecf5; }
.feg-ta-cal-day-today { font-weight: 700; color: #1a56db; }
.feg-ta-cal-day-selected { background: #1a56db !important; color: #fff !important; }
.feg-ta-selected-dates {
    display: flex;
    flex-wrap: wrap;
    gap: 5px;
    min-height: 26px;
    padding: 2px 0;
}
.feg-ta-date-tag {
    background: #e8f0fe;
    color: #1a56db;
    border: 1px solid #b3c9f9;
    border-radius: 12px;
    padding: 2px 10px;
    font-size: 12px;
    cursor: pointer;
    transition: background .1s, color .1s;
}
.feg-ta-date-tag:hover { background: #fde8e7; color: #c0392b; border-color: #f5c6c2; }
.feg-ta-preview {
    width: 100%;
    height: 160px;
    padding: 8px 10px;
    border: 1px solid #c9d1de;
    border-radius: 7px;
    font-size: 12px;
    font-family: "Cascadia Code", "Fira Code", Consolas, Monaco, monospace;
    resize: vertical;
    box-sizing: border-box;
    background: #f9fafb;
    color: #374151;
}
@media (max-width: 900px) {
    #feg-panel {
        right: 8px;
        bottom: 8px;
        max-width: calc(100vw - 16px);
    }
    .feg-header {
        gap: 4px;
    }
    .feg-header-btn {
        padding: 3px 6px;
    }
    .feg-tpl-btn {
        font-size: 12px;
    }
}
@media (prefers-reduced-motion: reduce) {
    .feg-body,
    .feg-tpl-btn {
        transition: none;
    }
}
`;
            document.head.appendChild(style);
        },

        _buildPanel() {
            clearTimeout(this._hoverCloseTimer);
            this._hoverCloseTimer = null;
            this._hovering = false;
            const panel = document.createElement('div');
            panel.id = PANEL_ID;
            panel.addEventListener('mouseenter', () => this._setPanelHovering(true));
            panel.addEventListener('mouseleave', () => this._schedulePanelHoverOff());

            // ── Header ──────────────────────────────────────────
            const header = document.createElement('div');
            header.className = 'feg-header';
            header.tabIndex = 0;
            header.setAttribute('role', 'button');

            const title = document.createElement('span');
            title.className = 'feg-header-title';
            title.textContent = PANEL_TRIGGER_TEXT;

            const hint = document.createElement('span');
            hint.className = 'feg-header-hint';
            hint.textContent = PANEL_ACTION_EXPAND;

            header.append(title, hint);
            header.addEventListener('click', () => this.togglePanelExpanded());
            header.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    this.togglePanelExpanded();
                }
            });

            // ── Body（模板按鈕）─────────────────────────────────
            const body = document.createElement('div');
            body.className = 'feg-body';

            const actionRow = document.createElement('div');
            actionRow.className = 'feg-action-row';

            const btnTemplate = document.createElement('button');
            btnTemplate.className = 'feg-header-btn feg-action-btn';
            btnTemplate.textContent = '⬇️';
            btnTemplate.title = 'Download example .xlsx template';
            btnTemplate.addEventListener('click', () => this._onDownloadTemplate());

            const isCustomize = this._templateMode === TEMPLATE_MODE_CUSTOMIZE;

            const btnPick = document.createElement('button');
            btnPick.className = 'feg-header-btn feg-action-btn';
            btnPick.textContent = '📂';
            btnPick.disabled = !isCustomize;
            btnPick.title = isCustomize
                ? 'Pick local Excel file (.xlsx) — saves a handle for later reload'
                : 'Switch to Customize mode to pick a local Excel file';
            btnPick.addEventListener('click', () => this._onPickExcel());

            const btnReload = document.createElement('button');
            btnReload.className = 'feg-header-btn feg-action-btn';
            btnReload.textContent = '🔄';
            btnReload.disabled = !isCustomize;
            btnReload.title = isCustomize
                ? 'Reload templates from the bound Excel file'
                : 'Switch to Customize mode to reload Excel templates';
            btnReload.addEventListener('click', () => this._onReloadExcel());

            const btnMode = document.createElement('button');
            btnMode.className = `feg-header-btn feg-action-btn feg-mode-toggle ${this._templateMode === TEMPLATE_MODE_DEFAULT ? 'feg-mode-default' : 'feg-mode-customize'}`;
            btnMode.textContent = this._templateMode === TEMPLATE_MODE_DEFAULT ? 'Default' : 'Cstm.';
            btnMode.title = `Switch to ${this._templateMode === TEMPLATE_MODE_DEFAULT ? 'Customize (Excel)' : 'Default'} mode`;
            btnMode.addEventListener('click', () => this._toggleTemplateMode());

            actionRow.append(btnTemplate, btnPick, btnReload, btnMode);
            body.appendChild(actionRow);

            // Customize 模式時顯示載入錯誤 / 提示
            if (this._templateMode === TEMPLATE_MODE_CUSTOMIZE && TemplateLoader.lastError) {
                const errBox = document.createElement('div');
                errBox.className = 'feg-empty-message';
                errBox.style.cssText = 'color:#b94a3a;white-space:pre-wrap;';
                errBox.textContent = '⚠ ' + TemplateLoader.lastError;
                body.appendChild(errBox);
            }

            if (this._templates.length === 0) {
                const empty = document.createElement('div');
                empty.className = 'feg-empty-message';
                empty.textContent = this._templateMode === TEMPLATE_MODE_CUSTOMIZE
                    ? '(No templates loaded — pick an Excel file or download the example template.)'
                    : '(No templates yet)';
                body.appendChild(empty);
            } else {
                // Sort: templates with `order` come first (ascending), those without
                // come last while preserving their original relative order (stable sort).
                const sorted = this._templates.slice().sort((a, b) => {
                    const aHasOrder = a.order != null;
                    const bHasOrder = b.order != null;
                    if (aHasOrder && bHasOrder) return a.order - b.order;
                    if (aHasOrder) return -1;
                    if (bHasOrder) return 1;
                    return 0;
                });
                sorted.forEach(tpl => {
                    const btn = this._buildTemplateButton(tpl);
                    body.appendChild(btn);
                });
            }

            if (this._travelApprovalConfig.enabled) {
                const taSep = document.createElement('div');
                taSep.className = 'feg-toolbar-sep';
                body.appendChild(taSep);

                const taBtn = document.createElement('button');
                taBtn.className = 'feg-tpl-btn';
                taBtn.style.background = this._travelApprovalConfig.button.color;
                taBtn.title = this._travelApprovalConfig.button.label;

                const taIcon = document.createElement('span');
                taIcon.className = 'feg-tpl-btn-icon';
                taIcon.textContent = this._travelApprovalConfig.button.icon;

                const taLabel = document.createElement('span');
                taLabel.textContent = this._travelApprovalConfig.button.label;

                taBtn.append(taIcon, taLabel);
                taBtn.addEventListener('click', () => this._onTravelApprovalClick());
                body.appendChild(taBtn);
            }

            panel.append(header, body);
            document.body.appendChild(panel);
            this._panel = panel;
            this._applyPanelVisibility();
            this._applyPanelExpandedState();
        },

        _buildTemplateButton(tpl) {
            const btn = document.createElement('button');
            btn.className = 'feg-tpl-btn';
            btn.style.background = (tpl.button && tpl.button.color) || '#555';
            btn.title = tpl.name || tpl.id;

            const icon = document.createElement('span');
            icon.className = 'feg-tpl-btn-icon';
            icon.textContent = (tpl.button && tpl.button.icon) || '✉';

            const label = document.createElement('span');
            label.textContent = (tpl.button && tpl.button.label) || tpl.name;

            btn.append(icon, label);
            btn.addEventListener('click', () => this._onTemplateClick(tpl));
            return btn;
        },

        _onTemplateClick(tpl) {
            try {
                switch (tpl.kind) {
                    case 'plain':     handlePlain(tpl);     break;
                    case 'external':  handleExternal(tpl);  break;
                    case 'scheduled': handleScheduled(tpl); break;
                    default:
                        this.toast(`Unknown kind: ${tpl.kind}`, 'error');
                }
            } catch (e) {
                this.toast(`Execution error: ${e.message}`, 'error');
                console.error('[FEG]', e);
            }
        },

        _onTravelApprovalClick() {
            try { TravelApprovalModal.show(this._travelApprovalConfig); }
            catch (e) {
                this.toast(`Execution error: ${e.message}`, 'error');
                console.error('[FEG]', e);
            }
        },

        _applyPanelVisibility() {
            if (!this._panel) return;
            this._panel.style.display = this._visible ? '' : 'none';
        },

        _applyPanelExpandedState() {
            if (!this._panel) return;
            this._panel.classList.toggle('feg-panel-open', !!this._expanded);
            const trigger = this._panel.querySelector('.feg-header');
            if (trigger) {
                const hint = trigger.querySelector('.feg-header-hint');
                const actionText = this._expanded ? PANEL_ACTION_COLLAPSE : PANEL_ACTION_EXPAND;
                trigger.setAttribute('aria-expanded', this._expanded ? 'true' : 'false');
                trigger.setAttribute('aria-label', `${actionText}${PANEL_ARIA_NAME}`);
                if (hint) hint.textContent = actionText;
            }
        },

        _setPanelHovering(isHovering) {
            if (!this._panel) return;
            if (isHovering) {
                clearTimeout(this._hoverCloseTimer);
                this._hoverCloseTimer = null;
            }
            this._hovering = !!isHovering;
            this._panel.classList.toggle('feg-panel-hovering', this._hovering);
        },

        _schedulePanelHoverOff() {
            clearTimeout(this._hoverCloseTimer);
            this._hoverCloseTimer = setTimeout(() => {
                if (!this._panel) return;
                if (this._panel.matches(':hover')) return;
                this._setPanelHovering(false);
            }, PANEL_HOVER_LEAVE_DELAY_MS);
        },

        _onDownloadTemplate() {
            try {
                ExcelTemplateBuilder.download();
                this.toast(`📤 Example ${TEMPLATE_FILE_NAME} downloaded.`, 'success');
            } catch (e) {
                this.toast(`Failed to build example workbook: ${e.message}`, 'error');
                console.error('[FEGX]', e);
            }
        },

        async _onPickExcel() {
            // Defensive guard: handler is also reachable via GM menu command,
            // so enforce the mode check regardless of button disabled state.
            if (this._templateMode !== TEMPLATE_MODE_CUSTOMIZE) {
                this.toast('ℹ️ Switch to Customize mode before picking an Excel file.', 'info');
                return;
            }
            if (!ExcelIO.isSupported()) {
                this.toast('File System Access API is not supported in this browser.', 'error');
                return;
            }
            let handle;
            try {
                handle = await ExcelIO.pickFile();
            } catch (e) {
                if (e && (e.name === 'AbortError' || /aborted/i.test(e.message || ''))) {
                    this.toast('File picker cancelled.', 'info');
                    return;
                }
                this.toast(`Failed to pick file: ${e.message}`, 'error');
                return;
            }
            try {
                await ExcelStore.saveHandle(handle);
            } catch (e) {
                this.toast(`Failed to save file handle: ${e.message}`, 'error');
                return;
            }
            const data = await TemplateLoader.loadTemplates({ canPrompt: true });
            this.reload(data);
            if (TemplateLoader.lastError) {
                this.toast(TemplateLoader.lastError, 'warn');
            } else {
                this.toast(`✅ Bound to "${handle.name}".`, 'success');
            }
        },

        async _onReloadExcel() {
            if (this._templateMode !== TEMPLATE_MODE_CUSTOMIZE) {
                this.toast('ℹ️ Switch to Customize mode before reloading Excel.', 'info');
                return;
            }
            const data = await TemplateLoader.loadTemplates({ canPrompt: true });
            this.reload(data);
            if (TemplateLoader.lastError) {
                this.toast(TemplateLoader.lastError, 'warn');
            }
        },

        async _onForgetExcel() {
            if (!confirm('Forget the bound Excel file? You will need to pick it again later.')) return;
            try {
                await ExcelStore.deleteHandle();
            } catch (e) {
                this.toast(`Failed to delete handle: ${e.message}`, 'error');
                return;
            }
            const data = await TemplateLoader.loadTemplates();
            this.reload(data);
            this.toast('🗑 Excel binding cleared.', 'success');
        },

        async _toggleTemplateMode() {
            const nextMode = this._templateMode === TEMPLATE_MODE_DEFAULT
                ? TEMPLATE_MODE_CUSTOMIZE
                : TEMPLATE_MODE_DEFAULT;
            Storage.saveTemplateMode(nextMode);
            // When switching INTO customize, allow permission prompt (we are inside a click handler)
            const data = await TemplateLoader.loadTemplates({ canPrompt: nextMode === TEMPLATE_MODE_CUSTOMIZE });
            this.reload(data);
            this._setPanelHovering(true);
            this._schedulePanelHoverOff();
        }
    };

    // ============================================================
    //  Toolbar popup commands (browser extension)
    //  Replaces the original Tampermonkey GM_registerMenuCommand entries.
    // ============================================================
    const COMMAND_HANDLERS = {
        downloadTemplate() {
            try {
                ExcelTemplateBuilder.download();
                UI.toast(`📤 Example ${TEMPLATE_FILE_NAME} downloaded.`, 'success');
            } catch (e) {
                UI.toast(`Failed: ${e.message}`, 'error');
            }
        },
        pickExcel() { UI._onPickExcel(); },
        reloadExcel() { UI._onReloadExcel(); },
        forgetExcel() { UI._onForgetExcel(); },
        togglePanel() { UI.togglePanelVisibility(); }
    };

    if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.onMessage) {
        chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
            if (!msg || msg.type !== 'feg.command') return false;
            const handler = COMMAND_HANDLERS[msg.command];
            if (!handler) {
                sendResponse({ ok: false, error: 'Unknown command: ' + msg.command });
                return false;
            }
            try {
                handler();
                sendResponse({ ok: true });
            } catch (e) {
                sendResponse({ ok: false, error: e && e.message ? e.message : String(e) });
            }
            return false;
        });
    }

    // ============================================================
    //  初始化（async：customize 模式可能要讀本地檔案）
    // ============================================================
    (async () => {
        try {
            // Wait for the GM-compat shim to finish loading persisted settings
            // from chrome.storage before any synchronous Storage.* call runs.
            if (typeof window !== 'undefined' && window.__FEG_INIT__) {
                try { await window.__FEG_INIT__; } catch (_) { /* ignore */ }
            }
            const initData = await TemplateLoader.loadTemplates();
            UI.init(initData);
            if (TemplateLoader.lastError) {
                console.warn('[FEGX] init:', TemplateLoader.lastError);
            }
        } catch (e) {
            console.error('[FEGX] init failed:', e);
            UI.init(EMPTY_DATA);
        }
    })();

})();
