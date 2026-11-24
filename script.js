/**
 * Automated Invoice & Receipt Parser for Accounting
 * Pure Client-Side Vanilla JavaScript (Gemini 2.5 Flash Vision + Math Audit + SheetJS + PDF.js)
 */

// Configure PDF.js worker
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

// ============================================================================
// 1. Bilingual Dictionary (Thai & English)
// ============================================================================
const I18N = {
  th: {
    appTitle: "ระบบแยกข้อมูลใบแจ้งหนี้และใบเสร็จอัตโนมัติเพื่องานบัญชี",
    appSubtitle: "Automated Invoice & Receipt Parser for Accounting (Gemini 2.5 Flash Vision)",
    wfStep1: "อัปโหลดใบเสร็จ (รูปภาพ/PDF)",
    wfStep2: "สกัดข้อมูลด้วย Gemini Vision",
    wfStep3: "ตรวจสอบยอดเงินและ VAT 7%",
    wfStep4: "ตรวจสอบและแก้ไขข้อมูลบัญชี",
    wfStep5: "ส่งออก Excel (.xlsx) / JSON",
    apiCardTitle: "การตั้งค่า Google Gemini API Key",
    apiStatusMissing: "ยังไม่ได้ใส่ Key",
    apiStatusReady: "พร้อมใช้งาน",
    apiCardDesc:
      "ประมวลผลบนเบราว์เซอร์ 100% คีย์ของคุณจะถูกเก็บไว้ในเครื่องของคุณเท่านั้น (ไม่ส่งผ่านเซิร์ฟเวอร์ตัวกลาง)",
    apiKeyPlaceholder: "กรอก Google Gemini API Key (AIza...)",
    saveKeyBtn: "💾 บันทึก Key ลงเครื่อง",
    clearKeyBtn: "🗑️ ลบ Key",
    uploadCardTitle: "อัปโหลดใบเสร็จ / ใบกำกับภาษี",
    dropzoneMain: "ลากไฟล์มาวางที่นี่ หรือ คลิกเพื่อเลือกไฟล์ (รองรับหลายไฟล์พร้อมกัน)",
    dropzoneSub: "รองรับใบเสร็จภาษาไทยและอังกฤษ (.png, .jpg, .jpeg, .pdf)",
    extractBtn: "✨ สกัดข้อมูลด้วย AI",
    loadDemoBtn: "🧪 โหลดตัวอย่างทดสอบ",
    clearAllBtn: "ล้างทั้งหมด",
    kpiTotalReceipts: "จำนวนใบเสร็จที่ประมวลผล",
    kpiReceiptsSub: "รายการทั้งหมดในระบบ",
    kpiTotalSpend: "ยอดใช้จ่ายสุทธิรวม (THB)",
    kpiSpendSub: "รวมภาษีมูลค่าเพิ่มแล้ว",
    kpiTotalVat: "ภาษีมูลค่าเพิ่ม 7% รวม",
    kpiVatSub: "ภาษีซื้อที่ตรวจสอบแล้ว",
    kpiPassRate: "อัตราความถูกต้องทางคณิตศาสตร์",
    tabInspector: "🔍 ตรวจสอบและแก้ไขใบเสร็จ",
    tabMetrics: "📊 ประสิทธิภาพระบบ & Ground Truth",
    exportExcelBtn: "📗 ส่งออก Excel (.xlsx)",
    exportJsonBtn: "📘 ส่งออก JSON (.json)",
    emptyTitle: "ยังไม่มีข้อมูลใบเสร็จที่สกัดแล้ว",
    emptyDesc:
      'อัปโหลดรูปภาพหรือไฟล์ PDF ใบเสร็จด้านบน แล้วกด "สกัดข้อมูลด้วย AI" หรือกด "โหลดตัวอย่างทดสอบ" เพื่อดูการทำงาน',
    selectReceiptLabel: "เลือกรายการใบเสร็จ:",
    previewTitle: "ภาพต้นฉบับเอกสาร",
    zoomBtn: "🔍 ขยายภาพ",
    headerFieldsTitle: "ข้อมูลหัวบิล / ร้านค้า (แก้ไขได้)",
    lblReceiptNo: "เลขที่ใบเสร็จ / ใบกำกับภาษี",
    lblMerchant: "ชื่อร้านค้า / บริษัท",
    lblDate: "วันที่ (DD/MM/YYYY)",
    lblTime: "เวลา (HH:MM)",
    lblTaxId: "เลขประจำตัวผู้เสียภาษี (13 หรือ 15 หลัก)",
    itemsTableTitle: "รายการสินค้า / บริการ",
    addItemBtn: "+ เพิ่มรายการ",
    thItemName: "รายการสินค้า",
    thQty: "จำนวน",
    thUnitPrice: "ราคา/หน่วย",
    thLineTotal: "รวมเงิน",
    thAction: "ลบ",
    totalsTitle: "สรุปยอดเงินและภาษีมูลค่าเพิ่ม 7%",
    recalcInclBtn: "คำนวณแบบรวม VAT 7%",
    recalcExclBtn: "คำนวณแบบแยก VAT 7%",
    lblSubtotal: "ยอดก่อนภาษี (Subtotal Excl. VAT)",
    lblVat: "ภาษีมูลค่าเพิ่ม 7% (VAT 7%)",
    lblGrandTotal: "ยอดสุทธิรวม (Grand Total)",
    timeBenchTitle: "เปรียบเทียบเวลาประมวลผล (AI vs. คีย์ข้อมูลด้วยมือ)",
    timeBenchDesc:
      "เปรียบเทียบความเร็วในการสกัดข้อมูลด้วย Gemini 2.5 Flash เทียบกับค่าเฉลี่ยการบันทึกบัญชีด้วยมนุษย์ (~90 วินาที/ใบเสร็จ)",
    aiTimeLabel: "🤖 เวลาประมวลผลด้วย AI (รวม)",
    manualTimeLabel: "🧑‍💻 เวลาประมาณการคีย์ด้วยมือ (90 วินาที/ใบ)",
    timeSavedLabel: "ประหยัดเวลาทำงานได้:",
    gtTitle: "ประเมินความแม่นยำเทียบเฉลย (Ground Truth Evaluator)",
    gtDesc:
      "อัปโหลดไฟล์ Excel เฉลย (.xlsx ที่มีชีต 'Receipt Summary' และ 'Receipt Details') เพื่อคำนวณความแม่นยำระดับฟิลด์ (Field-Level Accuracy %)",
    uploadGtBtn: "📂 อัปโหลดไฟล์เฉลย (.xlsx)",
    gtAccuracyLabel: "ความแม่นยำระดับฟิลด์โดยรวม (Field-Level Accuracy):",
    gtThReceipt: "เลขที่ใบเสร็จ",
    gtThMatched: "ฟิลด์ที่ตรงกัน",
    gtThAccuracy: "ความแม่นยำ (%)",
    gtThDiff: "รายละเอียดจุดต่าง",
    loadingTitle: "กำลังวิเคราะห์เอกสารด้วย Gemini 2.5 Flash...",

    // Dynamic Toast & Audit Messages
    toastKeySaved: "บันทึก API Key ลง localStorage เรียบร้อยแล้ว",
    toastKeyEmpty: "กรุณากรอก Google Gemini API Key ก่อนบันทึก",
    toastKeyCleared: "ลบ API Key ออกจากเครื่องแล้ว",
    toastInvalidFile: "ไม่รองรับไฟล์บางรายการ (รองรับเฉพาะ PNG, JPG, JPEG, PDF)",
    toastNoFiles: "กรุณาเลือกไฟล์ใบเสร็จก่อนเริ่มสกัดข้อมูล",
    toastNoApiKey: "กรุณากรอก Google Gemini API Key ก่อนใช้งาน AI",
    toastExtractSuccess: "สกัดข้อมูลใบเสร็จสำเร็จแล้ว!",
    toastDemoLoaded: "โหลดชุดข้อมูลใบเสร็จตัวอย่าง 2 รายการสำเร็จ (1 PASS, 1 MISMATCH)",
    toastExportExcel: "ส่งออกไฟล์ Excel (.xlsx) สำเร็จ",
    toastExportJson: "ส่งออกไฟล์ JSON (.json) สำเร็จ",
    toastGtNoData: "กรุณาสกัดข้อมูลใบเสร็จอย่างน้อย 1 รายการก่อนเทียบ Ground Truth",
    toastGtError: "ไม่สามารถอ่านไฟล์เฉลยได้ กรุณาตรวจสอบว่ามีชีต 'Receipt Summary'",
    errApiQuota: "เกินโควตาการเรียกใช้ API (429 Rate Limit) กรุณารอสักครู่แล้วลองใหม่",
    errApiInvalid: "API Key ไม่ถูกต้องหรือไม่มีสิทธิ์เข้าถึง (400/403)",
    errApiGeneral: "เกิดข้อผิดพลาดในการเชื่อมต่อ Gemini API: ",

    // Validation Audit Explanations
    auditPassSummary: "ตรวจสอบผ่าน: ยอดรวมรายการสินค้าและภาษีมูลค่าเพิ่ม 7% ถูกต้องตรงกัน",
    auditModeIncl: "รูปแบบใบเสร็จ: ราคารวมภาษีมูลค่าเพิ่มแล้ว (VAT-Inclusive)",
    auditModeExcl: "รูปแบบใบเสร็จ: ราคาแยกภาษีมูลค่าเพิ่ม (VAT-Exclusive)",
    auditItemSumMismatch:
      "ผลรวมรายการสินค้า ({itemsSum}) ไม่ตรงกับยอดก่อนภาษี ({sub}) และยอดสุทธิ ({grand})",
    auditVatSumMismatch:
      "ยอดก่อนภาษี ({sub}) + VAT 7% ({vat}) = {calcGrand} ไม่ตรงกับยอดสุทธิ ({grand}) [ผลต่าง {diff}]",
    auditRowMismatch:
      "แถวที่ {row} ({name}): {qty} × {price} = {expected} แต่ระบุรวมเงิน {actual}",
    auditVatRateNote:
      "หมายเหตุ: ยอด VAT ({vat}) ต่างจาก 7% ของยอดก่อนภาษี ({expectedVat}) เล็กน้อย"
  },

  en: {
    appTitle: "Automated Invoice & Receipt Parser for Accounting",
    appSubtitle: "End-to-End AI Accounting Automation (Gemini 2.5 Flash Vision)",
    wfStep1: "Upload Receipt (Image/PDF)",
    wfStep2: "Gemini Vision AI Extraction",
    wfStep3: "Math & 7% VAT Validation",
    wfStep4: "Interactive Audit & Editor",
    wfStep5: "Export Excel (.xlsx) / JSON",
    apiCardTitle: "Google Gemini API Key Setup",
    apiStatusMissing: "Key Required",
    apiStatusReady: "Ready",
    apiCardDesc:
      "100% client-side execution. Your API key stays strictly in your browser storage and is never sent to any third-party backend.",
    apiKeyPlaceholder: "Enter Google Gemini API Key (AIza...)",
    saveKeyBtn: "💾 Save Key Locally",
    clearKeyBtn: "🗑️ Clear Key",
    uploadCardTitle: "Upload Receipts / Tax Invoices",
    dropzoneMain: "Drag & drop files here, or click to browse (Batch upload supported)",
    dropzoneSub: "Supports Thai & English receipts (.png, .jpg, .jpeg, .pdf)",
    extractBtn: "✨ Extract with AI",
    loadDemoBtn: "🧪 Load Demo Sample",
    clearAllBtn: "Clear All",
    kpiTotalReceipts: "Total Receipts Processed",
    kpiReceiptsSub: "Verified records in session",
    kpiTotalSpend: "Total Net Spend (THB)",
    kpiSpendSub: "Inclusive of 7% VAT",
    kpiTotalVat: "Total VAT 7%",
    kpiVatSub: "Audited input tax",
    kpiPassRate: "Validation Pass Rate",
    tabInspector: "🔍 Receipt Inspector & Editor",
    tabMetrics: "📊 System Metrics & Ground Truth",
    exportExcelBtn: "📗 Export Excel (.xlsx)",
    exportJsonBtn: "📘 Export JSON (.json)",
    emptyTitle: "No Extracted Receipts Yet",
    emptyDesc:
      'Upload receipt images or PDFs above and click "Extract with AI", or click "Load Demo Sample" to explore the workspace.',
    selectReceiptLabel: "Select Receipt:",
    previewTitle: "Original Document Preview",
    zoomBtn: "🔍 Zoom",
    headerFieldsTitle: "Header & Merchant Metadata (Editable)",
    lblReceiptNo: "Receipt / Tax Invoice No.",
    lblMerchant: "Merchant / Company Name",
    lblDate: "Date (DD/MM/YYYY)",
    lblTime: "Time (HH:MM)",
    lblTaxId: "Tax ID / Biller ID (13 or 15 digits)",
    itemsTableTitle: "Line Items",
    addItemBtn: "+ Add Row",
    thItemName: "Item Description",
    thQty: "Qty",
    thUnitPrice: "Unit Price",
    thLineTotal: "Line Total",
    thAction: "Del",
    totalsTitle: "Totals & 7% VAT Summary",
    recalcInclBtn: "Recalc (VAT-Inclusive)",
    recalcExclBtn: "Recalc (VAT-Exclusive)",
    lblSubtotal: "Subtotal (Excl. VAT)",
    lblVat: "VAT 7% Amount",
    lblGrandTotal: "Grand Total Paid",
    timeBenchTitle: "Processing Time Benchmark (AI vs. Manual Entry)",
    timeBenchDesc:
      "Compares Gemini 2.5 Flash extraction speed against the ~90 seconds/receipt manual accounting entry benchmark.",
    aiTimeLabel: "🤖 Total AI Extraction Time",
    manualTimeLabel: "🧑‍💻 Estimated Manual Entry Time (90s/receipt)",
    timeSavedLabel: "Total Time Saved:",
    gtTitle: "Ground Truth Accuracy Evaluator",
    gtDesc:
      "Upload a reference Excel file (.xlsx with 'Receipt Summary' and 'Receipt Details' sheets) to compute field-level extraction accuracy (%).",
    uploadGtBtn: "📂 Upload Reference (.xlsx)",
    gtAccuracyLabel: "Overall Field-Level Extraction Accuracy:",
    gtThReceipt: "Receipt No.",
    gtThMatched: "Matched Fields",
    gtThAccuracy: "Accuracy (%)",
    gtThDiff: "Discrepancy Notes",
    loadingTitle: "Extracting structured data via Gemini 2.5 Flash...",

    // Dynamic Toast & Audit Messages
    toastKeySaved: "API Key saved to localStorage.",
    toastKeyEmpty: "Please enter a valid Google Gemini API Key first.",
    toastKeyCleared: "API Key removed from localStorage.",
    toastInvalidFile: "Some files were skipped (only PNG, JPG, JPEG, PDF allowed).",
    toastNoFiles: "Please select at least one receipt file first.",
    toastNoApiKey: "Please enter your Google Gemini API Key before extracting.",
    toastExtractSuccess: "Receipt extraction completed!",
    toastDemoLoaded: "Loaded 2 sample receipts (1 PASS, 1 MISMATCH for audit demo).",
    toastExportExcel: "Exported Excel workbook (.xlsx) successfully.",
    toastExportJson: "Exported structured JSON (.json) successfully.",
    toastGtNoData: "Please extract at least 1 receipt before running Ground Truth evaluation.",
    toastGtError: "Could not parse reference file. Ensure 'Receipt Summary' sheet exists.",
    errApiQuota: "API quota exceeded (429 Rate Limit). Please wait a moment and retry.",
    errApiInvalid: "Invalid or unauthorized API Key (400/403).",
    errApiGeneral: "Gemini API request failed: ",

    // Validation Audit Explanations
    auditPassSummary: "Validation Passed: Line-item totals and 7% VAT reconcile accurately.",
    auditModeIncl: "Tax Mode Detected: VAT-Inclusive Pricing",
    auditModeExcl: "Tax Mode Detected: VAT-Exclusive Pricing",
    auditItemSumMismatch:
      "Sum of items ({itemsSum}) matches neither Subtotal ({sub}) nor Grand Total ({grand})",
    auditVatSumMismatch:
      "Subtotal ({sub}) + VAT 7% ({vat}) = {calcGrand}, which differs from Grand Total ({grand}) [diff: {diff}]",
    auditRowMismatch:
      "Row {row} ({name}): {qty} × {price} = {expected}, but line total shows {actual}",
    auditVatRateNote:
      "Note: Reported VAT ({vat}) differs slightly from 7% of Subtotal ({expectedVat})"
  }
};

// ============================================================================
// 2. Application State
// ============================================================================
const STORAGE_KEY = "gemini_accounting_api_key";

const state = {
  lang: "th",
  apiKey: localStorage.getItem(STORAGE_KEY) || "",
  pendingFiles: [], // File objects waiting to be processed
  receipts: [], // Processed receipt records
  activeReceiptIndex: 0,
  pdfDoc: null,
  pdfPageNum: 1,
  groundTruthResults: null
};

// ============================================================================
// 3. DOM References
// ============================================================================
const el = {
  langThBtn: document.getElementById("langThBtn"),
  langEnBtn: document.getElementById("langEnBtn"),
  workflowSteps: document.getElementById("workflowSteps"),

  apiKeyInput: document.getElementById("apiKeyInput"),
  toggleApiVisibilityBtn: document.getElementById("toggleApiVisibilityBtn"),
  saveApiKeyBtn: document.getElementById("saveApiKeyBtn"),
  clearApiKeyBtn: document.getElementById("clearApiKeyBtn"),
  apiStatusBadge: document.getElementById("apiStatusBadge"),

  dropZone: document.getElementById("dropZone"),
  fileInput: document.getElementById("fileInput"),
  pendingFilesContainer: document.getElementById("pendingFilesContainer"),
  processBtn: document.getElementById("processBtn"),
  loadDemoBtn: document.getElementById("loadDemoBtn"),
  clearAllBtn: document.getElementById("clearAllBtn"),

  kpiReceiptsCount: document.getElementById("kpiReceiptsCount"),
  kpiNetSpend: document.getElementById("kpiNetSpend"),
  kpiTotalVat: document.getElementById("kpiTotalVat"),
  kpiPassRate: document.getElementById("kpiPassRate"),
  kpiPassCount: document.getElementById("kpiPassCount"),

  tabInspectorBtn: document.getElementById("tabInspectorBtn"),
  tabMetricsBtn: document.getElementById("tabMetricsBtn"),
  inspectorView: document.getElementById("inspectorView"),
  metricsView: document.getElementById("metricsView"),

  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportJsonBtn: document.getElementById("exportJsonBtn"),

  emptyState: document.getElementById("emptyState"),
  activeInspector: document.getElementById("activeInspector"),
  receiptSelectorList: document.getElementById("receiptSelectorList"),

  receiptImagePreview: document.getElementById("receiptImagePreview"),
  receiptPdfCanvas: document.getElementById("receiptPdfCanvas"),
  demoPreviewPlaceholder: document.getElementById("demoPreviewPlaceholder"),
  previewFileMeta: document.getElementById("previewFileMeta"),
  pdfPrevBtn: document.getElementById("pdfPrevBtn"),
  pdfNextBtn: document.getElementById("pdfNextBtn"),
  pdfPageInfo: document.getElementById("pdfPageInfo"),
  zoomToggleBtn: document.getElementById("zoomToggleBtn"),

  validationCard: document.getElementById("validationCard"),
  validationBadge: document.getElementById("validationBadge"),
  extractionTimeBadge: document.getElementById("extractionTimeBadge"),
  validationDetails: document.getElementById("validationDetails"),

  fieldReceiptNo: document.getElementById("fieldReceiptNo"),
  fieldMerchant: document.getElementById("fieldMerchant"),
  fieldDate: document.getElementById("fieldDate"),
  fieldTime: document.getElementById("fieldTime"),
  fieldTaxId: document.getElementById("fieldTaxId"),

  itemsTableBody: document.getElementById("itemsTableBody"),
  addItemBtn: document.getElementById("addItemBtn"),

  fieldSubtotal: document.getElementById("fieldSubtotal"),
  fieldVat: document.getElementById("fieldVat"),
  fieldGrandTotal: document.getElementById("fieldGrandTotal"),
  recalcInclBtn: document.getElementById("recalcInclBtn"),
  recalcExclBtn: document.getElementById("recalcExclBtn"),

  metricAiTime: document.getElementById("metricAiTime"),
  metricManualTime: document.getElementById("metricManualTime"),
  aiTimeBar: document.getElementById("aiTimeBar"),
  manualTimeBar: document.getElementById("manualTimeBar"),
  metricTimeSavedPct: document.getElementById("metricTimeSavedPct"),
  metricTimeSavedSec: document.getElementById("metricTimeSavedSec"),

  groundTruthInput: document.getElementById("groundTruthInput"),
  gtResultBox: document.getElementById("gtResultBox"),
  gtAccuracyValue: document.getElementById("gtAccuracyValue"),
  gtTableBody: document.getElementById("gtTableBody"),

  zoomModal: document.getElementById("zoomModal"),
  zoomedImg: document.getElementById("zoomedImg"),
  closeZoomBtn: document.getElementById("closeZoomBtn"),

  loadingOverlay: document.getElementById("loadingOverlay"),
  loadingProgress: document.getElementById("loadingProgress"),
  toastContainer: document.getElementById("toastContainer")
};

// ============================================================================
// 4. Initialization & Event Listeners
// ============================================================================
function init() {
  if (state.apiKey) {
    el.apiKeyInput.value = state.apiKey;
  }
  updateApiStatusBadge();
  applyLanguage(state.lang);
  setWorkflowStage(1);
  bindEvents();
}

function bindEvents() {
  // Language Switcher
  el.langThBtn.addEventListener("click", () => applyLanguage("th"));
  el.langEnBtn.addEventListener("click", () => applyLanguage("en"));

  // API Key Controls
  el.toggleApiVisibilityBtn.addEventListener("click", () => {
    el.apiKeyInput.type = el.apiKeyInput.type === "password" ? "text" : "password";
  });

  el.apiKeyInput.addEventListener("input", (e) => {
    state.apiKey = e.target.value.trim();
    updateApiStatusBadge();
  });

  el.saveApiKeyBtn.addEventListener("click", () => {
    const val = el.apiKeyInput.value.trim();
    if (!val) {
      showToast(t("toastKeyEmpty"), "error");
      return;
    }
    state.apiKey = val;
    localStorage.setItem(STORAGE_KEY, val);
    updateApiStatusBadge();
    showToast(t("toastKeySaved"), "success");
  });

  el.clearApiKeyBtn.addEventListener("click", () => {
    state.apiKey = "";
    el.apiKeyInput.value = "";
    localStorage.removeItem(STORAGE_KEY);
    updateApiStatusBadge();
    showToast(t("toastKeyCleared"));
  });

  // File Dropzone & Input
  el.dropZone.addEventListener("click", () => el.fileInput.click());
  el.dropZone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      el.fileInput.click();
    }
  });

  el.dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    el.dropZone.classList.add("dragover");
  });

  el.dropZone.addEventListener("dragleave", () => {
    el.dropZone.classList.remove("dragover");
  });

  el.dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    el.dropZone.classList.remove("dragover");
    handleSelectedFiles(e.dataTransfer.files);
  });

  el.fileInput.addEventListener("change", (e) => {
    handleSelectedFiles(e.target.files);
    el.fileInput.value = "";
  });

  // Action Buttons
  el.processBtn.addEventListener("click", processBatchReceipts);
  el.loadDemoBtn.addEventListener("click", loadDemoDataset);
  el.clearAllBtn.addEventListener("click", clearAllData);

  // Navigation View Tabs
  [el.tabInspectorBtn, el.tabMetricsBtn].forEach((tab) => {
    tab.addEventListener("click", () => {
      el.tabInspectorBtn.classList.remove("active");
      el.tabMetricsBtn.classList.remove("active");
      el.tabInspectorBtn.setAttribute("aria-selected", "false");
      el.tabMetricsBtn.setAttribute("aria-selected", "false");

      tab.classList.add("active");
      tab.setAttribute("aria-selected", "true");

      const targetId = tab.dataset.target;
      el.inspectorView.classList.toggle("hidden", targetId !== "inspectorView");
      el.metricsView.classList.toggle("hidden", targetId !== "metricsView");
    });
  });

  // Export Buttons
  el.exportExcelBtn.addEventListener("click", exportToExcel);
  el.exportJsonBtn.addEventListener("click", exportToJson);

  // PDF Navigation & Image Zoom
  el.pdfPrevBtn.addEventListener("click", () => changePdfPage(-1));
  el.pdfNextBtn.addEventListener("click", () => changePdfPage(1));
  el.zoomToggleBtn.addEventListener("click", openZoomModal);
  el.closeZoomBtn.addEventListener("click", () => el.zoomModal.classList.add("hidden"));
  el.zoomModal.addEventListener("click", (e) => {
    if (e.target === el.zoomModal) el.zoomModal.classList.add("hidden");
  });

  // Editable Header & Totals Inputs
  const boundInputs = [
    el.fieldReceiptNo,
    el.fieldMerchant,
    el.fieldDate,
    el.fieldTime,
    el.fieldTaxId,
    el.fieldSubtotal,
    el.fieldVat,
    el.fieldGrandTotal
  ];

  boundInputs.forEach((input) => {
    input.addEventListener("input", () => {
      const current = getCurrentReceipt();
      if (!current) return;
      const field = input.dataset.field;
      if (["subtotal_ex_vat", "vat_7_percent", "grand_total"].includes(field)) {
        current.data[field] = parseFloat(input.value) || 0;
      } else {
        current.data[field] = input.value;
      }
      runAuditAndRefreshUI();
    });
  });

  // Add Item & Recalculation Helpers
  el.addItemBtn.addEventListener("click", () => {
    const current = getCurrentReceipt();
    if (!current) return;
    current.data.items.push({
      item_name: state.lang === "th" ? "รายการใหม่" : "New Item",
      qty: 1,
      unit_price: 0,
      line_total: 0
    });
    renderItemsTable(current.data.items);
    runAuditAndRefreshUI();
  });

  el.recalcInclBtn.addEventListener("click", () => recalculateTotals("inclusive"));
  el.recalcExclBtn.addEventListener("click", () => recalculateTotals("exclusive"));

  // Ground Truth Upload
  el.groundTruthInput.addEventListener("change", handleGroundTruthUpload);
}

// ============================================================================
// 5. Bilingual Localization Helper
// ============================================================================
function t(key) {
  return I18N[state.lang][key] || key;
}

function applyLanguage(lang) {
  state.lang = lang;
  document.documentElement.lang = lang;

  el.langThBtn.classList.toggle("active", lang === "th");
  el.langEnBtn.classList.toggle("active", lang === "en");

  // Translate all static [data-i18n] elements
  document.querySelectorAll("[data-i18n]").forEach((node) => {
    const key = node.getAttribute("data-i18n");
    if (I18N[lang][key]) {
      node.textContent = I18N[lang][key];
    }
  });

  // Translate placeholders
  document.querySelectorAll("[data-i18n-placeholder]").forEach((node) => {
    const key = node.getAttribute("data-i18n-placeholder");
    if (I18N[lang][key]) {
      node.placeholder = I18N[lang][key];
    }
  });

  updateApiStatusBadge();

  // Re-render dynamic sections without losing state
  if (state.receipts.length > 0) {
    renderReceiptSelector();
    renderActiveReceipt();
    updateDashboardMetrics();
  }
  if (state.groundTruthResults) {
    renderGroundTruthTable();
  }
}

function updateApiStatusBadge() {
  const hasKey = Boolean(state.apiKey && state.apiKey.length > 10);
  el.apiStatusBadge.textContent = hasKey ? t("apiStatusReady") : t("apiStatusMissing");
  el.apiStatusBadge.className = `status-pill ${hasKey ? "status-ok" : "status-warn"}`;
}

function setWorkflowStage(activeStep) {
  const steps = el.workflowSteps.querySelectorAll(".wf-step");
  steps.forEach((stepEl) => {
    const num = Number(stepEl.dataset.step);
    stepEl.classList.remove("active", "done");
    if (num < activeStep) stepEl.classList.add("done");
    if (num === activeStep) stepEl.classList.add("active");
  });
}

// ============================================================================
// 6. File Selection & Base64 Conversion
// ============================================================================
function handleSelectedFiles(fileList) {
  const allowedTypes = ["image/png", "image/jpeg", "application/pdf"];
  let invalidCount = 0;

  Array.from(fileList).forEach((file) => {
    if (allowedTypes.includes(file.type)) {
      state.pendingFiles.push(file);
    } else {
      invalidCount++;
    }
  });

  if (invalidCount > 0) {
    showToast(t("toastInvalidFile"), "error");
  }

  renderPendingFiles();
}

function renderPendingFiles() {
  if (state.pendingFiles.length === 0) {
    el.pendingFilesContainer.classList.add("hidden");
    el.pendingFilesContainer.innerHTML = "";
    el.processBtn.disabled = true;
    return;
  }

  el.pendingFilesContainer.classList.remove("hidden");
  el.processBtn.disabled = false;

  el.pendingFilesContainer.innerHTML = state.pendingFiles
    .map(
      (f, idx) => `
      <span class="file-chip">
        <span>${f.type === "application/pdf" ? "📑" : "🖼️"} ${escapeHtml(f.name)}</span>
        <button type="button" onclick="removePendingFile(${idx})" style="border:none;background:none;cursor:pointer;font-weight:700;color:inherit;">×</button>
      </span>
    `
    )
    .join("");
}

window.removePendingFile = function (index) {
  state.pendingFiles.splice(index, 1);
  renderPendingFiles();
};

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      const base64Data = dataUrl.split(",")[1];
      resolve({ base64Data, dataUrl });
    };
    reader.onerror = (err) => reject(err);
    reader.readAsDataURL(file);
  });
}

// ============================================================================
// 7. Google Gemini 2.5 Flash Vision REST API Integration
// ============================================================================
const GEMINI_RESPONSE_SCHEMA = {
  type: "OBJECT",
  properties: {
    receipt_no: {
      type: "STRING",
      description: "Invoice or Receipt number, e.g., INV-20250816-001. Empty string if missing."
    },
    date: {
      type: "STRING",
      description:
        "Standardized date in DD/MM/YYYY format (Gregorian CE year; convert Buddhist Era year like 2568/2569 to 2025/2026)."
    },
    time: {
      type: "STRING",
      description: "Transaction time in HH:MM 24-hour format, or empty string if unavailable."
    },
    merchant_name: {
      type: "STRING",
      description: "Store, restaurant, or company name."
    },
    tax_id: {
      type: "STRING",
      description: "13-digit or 15-digit Tax ID / Biller ID (digits only)."
    },
    items: {
      type: "ARRAY",
      description: "List of purchased line items.",
      items: {
        type: "OBJECT",
        properties: {
          item_name: { type: "STRING" },
          qty: { type: "NUMBER" },
          unit_price: { type: "NUMBER" },
          line_total: { type: "NUMBER" }
        },
        required: ["item_name", "qty", "unit_price", "line_total"]
      }
    },
    subtotal_ex_vat: {
      type: "NUMBER",
      description: "Subtotal amount before 7% VAT. If receipt only shows VAT-inclusive total, compute or extract pre-VAT subtotal."
    },
    vat_7_percent: {
      type: "NUMBER",
      description: "7% VAT amount in THB. Use 0 if receipt is VAT-exempt."
    },
    grand_total: {
      type: "NUMBER",
      description: "Final net total paid."
    }
  },
  required: [
    "receipt_no",
    "date",
    "time",
    "merchant_name",
    "tax_id",
    "items",
    "subtotal_ex_vat",
    "vat_7_percent",
    "grand_total"
  ]
};

async function callGeminiVisionAPI(base64Data, mimeType, apiKey) {
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(
    apiKey
  )}`;

  const systemPrompt = `You are an expert Thai & International Certified Public Accountant (CPA) OCR and data extraction engine.
Analyze the attached receipt or tax invoice image/PDF and extract all fields strictly matching the JSON schema:
1. Standardize 'date' to DD/MM/YYYY (convert Thai Buddhist Era years such as 2568 or 2569 into Gregorian CE years 2025 or 2026).
2. Standardize 'time' to HH:MM (24-hour format).
3. Extract 'tax_id' as a clean 13-digit Tax ID (เลขประจำตัวผู้เสียภาษี) or 15-digit Biller ID.
4. For each line item in 'items', extract 'item_name', numeric 'qty', numeric 'unit_price', and numeric 'line_total'.
5. Extract 'subtotal_ex_vat' (amount before 7% VAT), 'vat_7_percent', and 'grand_total' (final net total). If a Thai retail receipt shows VAT-included prices and prints a breakdown at the bottom (e.g., Vatable / Before VAT and VAT 7%), use those exact printed figures.`;

  const payload = {
    contents: [
      {
        role: "user",
        parts: [
          { text: systemPrompt },
          {
            inlineData: {
              mimeType,
              data: base64Data
            }
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json",
      responseSchema: GEMINI_RESPONSE_SCHEMA
    }
  };

  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    if (response.status === 429) throw new Error(t("errApiQuota"));
    if (response.status === 400 || response.status === 403) throw new Error(t("errApiInvalid"));
    const errText = await response.text();
    throw new Error(`${t("errApiGeneral")}${response.status} ${errText}`);
  }

  const result = await response.json();
  const rawText = result?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!rawText) {
    throw new Error("Empty structured response from Gemini Vision API.");
  }

  return JSON.parse(rawText);
}

async function processBatchReceipts() {
  if (state.pendingFiles.length === 0) {
    showToast(t("toastNoFiles"), "error");
    return;
  }
  if (!state.apiKey) {
    showToast(t("toastNoApiKey"), "error");
    el.apiKeyInput.focus();
    return;
  }

  setWorkflowStage(2);
  el.loadingOverlay.classList.remove("hidden");

  const filesToProcess = [...state.pendingFiles];
  let successCount = 0;

  for (let i = 0; i < filesToProcess.length; i++) {
    const file = filesToProcess[i];
    el.loadingProgress.textContent = `${i + 1} / ${filesToProcess.length} — ${file.name}`;

    try {
      const startTime = performance.now();
      const { base64Data, dataUrl } = await fileToBase64(file);
      const extracted = await callGeminiVisionAPI(base64Data, file.type, state.apiKey);
      const elapsedSeconds = Number(((performance.now() - startTime) / 1000).toFixed(2));

      const normalized = normalizeExtractedData(extracted);
      const audit = validateReceiptMath(normalized);

      state.receipts.push({
        id: `rec_${Date.now()}_${i}`,
        fileName: file.name,
        mimeType: file.type,
        dataUrl,
        base64Data,
        extractionTimeSec: elapsedSeconds,
        data: normalized,
        audit
      });

      successCount++;
    } catch (err) {
      console.error("Extraction error:", err);
      showToast(`${file.name}: ${err.message}`, "error");
    }
  }

  el.loadingOverlay.classList.add("hidden");

  if (successCount > 0) {
    state.pendingFiles = [];
    renderPendingFiles();
    state.activeReceiptIndex = state.receipts.length - 1;
    setWorkflowStage(4);
    refreshFullWorkspace();
    showToast(`${t("toastExtractSuccess")} (${successCount})`, "success");
  } else {
    setWorkflowStage(1);
  }
}

function normalizeExtractedData(raw) {
  return {
    receipt_no: String(raw.receipt_no || "").trim(),
    date: String(raw.date || "").trim(),
    time: String(raw.time || "").trim(),
    merchant_name: String(raw.merchant_name || "").trim(),
    tax_id: String(raw.tax_id || "").trim(),
    items: Array.isArray(raw.items)
      ? raw.items.map((item) => ({
          item_name: String(item.item_name || ""),
          qty: Number(item.qty) || 0,
          unit_price: Number(item.unit_price) || 0,
          line_total: Number(item.line_total) || 0
        }))
      : [],
    subtotal_ex_vat: Number(raw.subtotal_ex_vat) || 0,
    vat_7_percent: Number(raw.vat_7_percent) || 0,
    grand_total: Number(raw.grand_total) || 0
  };
}

// ============================================================================
// 8. Automated Math & Accounting Validation Engine
// ============================================================================
function validateReceiptMath(data) {
  const TOLERANCE = 0.05;
  const issues = [];
  const notes = [];

  // 1. Check each row's qty * unit_price vs line_total & sum qty * unit_price
  let calculatedItemsSum = 0;
  data.items.forEach((item, idx) => {
    const rowExpected = round2(item.qty * item.unit_price);
    calculatedItemsSum += rowExpected;

    if (Math.abs(rowExpected - item.line_total) > TOLERANCE) {
      issues.push({
        key: "auditRowMismatch",
        params: {
          row: idx + 1,
          name: item.item_name || `#${idx + 1}`,
          qty: item.qty,
          price: item.unit_price.toFixed(2),
          expected: rowExpected.toFixed(2),
          actual: item.line_total.toFixed(2)
        }
      });
    }
  });
  calculatedItemsSum = round2(calculatedItemsSum);

  // 2. Line-Item Sum Check against Subtotal (VAT-Exclusive) or Grand Total (VAT-Inclusive)
  const matchesSubtotal = Math.abs(calculatedItemsSum - data.subtotal_ex_vat) <= TOLERANCE;
  const matchesGrandTotal = Math.abs(calculatedItemsSum - data.grand_total) <= TOLERANCE;

  if (!matchesSubtotal && !matchesGrandTotal) {
    issues.push({
      key: "auditItemSumMismatch",
      params: {
        itemsSum: calculatedItemsSum.toFixed(2),
        sub: data.subtotal_ex_vat.toFixed(2),
        grand: data.grand_total.toFixed(2)
      }
    });
  } else {
    notes.push({
      key: matchesGrandTotal && !matchesSubtotal ? "auditModeIncl" : "auditModeExcl",
      params: {}
    });
  }

  // 3. VAT 7% Consistency Check: subtotal_ex_vat + vat_7_percent == grand_total (±0.05)
  const expectedGrandTotal = round2(data.subtotal_ex_vat + data.vat_7_percent);
  const grandDiff = round2(data.grand_total - expectedGrandTotal);

  if (Math.abs(grandDiff) > TOLERANCE) {
    issues.push({
      key: "auditVatSumMismatch",
      params: {
        sub: data.subtotal_ex_vat.toFixed(2),
        vat: data.vat_7_percent.toFixed(2),
        calcGrand: expectedGrandTotal.toFixed(2),
        grand: data.grand_total.toFixed(2),
        diff: (grandDiff > 0 ? "+" : "") + grandDiff.toFixed(2)
      }
    });
  }

  // Soft informational note if VAT > 0 and differs from exact 7% of subtotal by > 0.15
  if (data.vat_7_percent > 0) {
    const expected7Pct = round2(data.subtotal_ex_vat * 0.07);
    if (Math.abs(expected7Pct - data.vat_7_percent) > 0.15) {
      notes.push({
        key: "auditVatRateNote",
        params: {
          vat: data.vat_7_percent.toFixed(2),
          expectedVat: expected7Pct.toFixed(2)
        }
      });
    }
  }

  const isPass = issues.length === 0;
  return {
    status: isPass ? "✅ PASS" : "⚠️ MISMATCH",
    isPass,
    calculatedItemsSum,
    issues,
    notes
  };
}

function formatAuditMsg(entry) {
  let msg = t(entry.key);
  Object.entries(entry.params || {}).forEach(([k, v]) => {
    msg = msg.replaceAll(`{${k}}`, v);
  });
  return msg;
}

function round2(num) {
  return Math.round((Number(num) + Number.EPSILON) * 100) / 100;
}

// ============================================================================
// 9. Interactive Workspace Rendering & Editing
// ============================================================================
function getCurrentReceipt() {
  return state.receipts[state.activeReceiptIndex] || null;
}

function refreshFullWorkspace() {
  const hasReceipts = state.receipts.length > 0;
  el.emptyState.classList.toggle("hidden", hasReceipts);
  el.activeInspector.classList.toggle("hidden", !hasReceipts);
  el.exportExcelBtn.disabled = !hasReceipts;
  el.exportJsonBtn.disabled = !hasReceipts;

  if (hasReceipts) {
    renderReceiptSelector();
    renderActiveReceipt();
  }
  updateDashboardMetrics();
}

function renderReceiptSelector() {
  el.receiptSelectorList.innerHTML = state.receipts
    .map((rec, idx) => {
      const label = rec.data.receipt_no || rec.fileName;
      const icon = rec.audit.isPass ? "✅" : "⚠️";
      const activeCls = idx === state.activeReceiptIndex ? "active" : "";
      return `
        <button type="button" class="receipt-pill ${activeCls}" onclick="selectReceipt(${idx})">
          <span>${icon}</span>
          <span>${escapeHtml(label)}</span>
        </button>
      `;
    })
    .join("");
}

window.selectReceipt = function (index) {
  state.activeReceiptIndex = index;
  renderReceiptSelector();
  renderActiveReceipt();
};

async function renderActiveReceipt() {
  const rec = getCurrentReceipt();
  if (!rec) return;

  // 1. Render Document Preview
  await renderDocumentPreview(rec);

  // 2. Populate Header Fields
  el.fieldReceiptNo.value = rec.data.receipt_no;
  el.fieldMerchant.value = rec.data.merchant_name;
  el.fieldDate.value = rec.data.date;
  el.fieldTime.value = rec.data.time;
  el.fieldTaxId.value = rec.data.tax_id;

  // 3. Populate Totals
  el.fieldSubtotal.value = rec.data.subtotal_ex_vat.toFixed(2);
  el.fieldVat.value = rec.data.vat_7_percent.toFixed(2);
  el.fieldGrandTotal.value = rec.data.grand_total.toFixed(2);

  // 4. Populate Line Items Table
  renderItemsTable(rec.data.items);

  // 5. Render Validation Status Banner
  renderValidationBanner(rec);
}

async function renderDocumentPreview(rec) {
  el.receiptImagePreview.classList.add("hidden");
  el.receiptPdfCanvas.classList.add("hidden");
  el.demoPreviewPlaceholder.classList.add("hidden");
  el.pdfPrevBtn.classList.add("hidden");
  el.pdfNextBtn.classList.add("hidden");
  el.pdfPageInfo.classList.add("hidden");
  el.zoomToggleBtn.classList.add("hidden");

  el.previewFileMeta.textContent = `${rec.fileName} (${rec.mimeType})`;

  if (rec.isDemo) {
    el.demoPreviewPlaceholder.classList.remove("hidden");
    el.demoPreviewPlaceholder.innerHTML = rec.demoHtml;
    return;
  }

  if (rec.mimeType === "application/pdf" && window.pdfjsLib) {
    try {
      const rawBytes = atob(rec.base64Data);
      const uint8Array = new Uint8Array(rawBytes.length);
      for (let i = 0; i < rawBytes.length; i++) {
        uint8Array[i] = rawBytes.charCodeAt(i);
      }
      state.pdfDoc = await pdfjsLib.getDocument({ data: uint8Array }).promise;
      state.pdfPageNum = 1;
      el.receiptPdfCanvas.classList.remove("hidden");
      if (state.pdfDoc.numPages > 1) {
        el.pdfPrevBtn.classList.remove("hidden");
        el.pdfNextBtn.classList.remove("hidden");
        el.pdfPageInfo.classList.remove("hidden");
      }
      await renderPdfPage(state.pdfPageNum);
    } catch (err) {
      console.error("PDF preview render error:", err);
    }
  } else {
    el.receiptImagePreview.src = rec.dataUrl;
    el.receiptImagePreview.classList.remove("hidden");
    el.zoomToggleBtn.classList.remove("hidden");
  }
}

async function renderPdfPage(pageNum) {
  if (!state.pdfDoc) return;
  const page = await state.pdfDoc.getPage(pageNum);
  const viewport = page.getViewport({ scale: 1.35 });
  const canvas = el.receiptPdfCanvas;
  const ctx = canvas.getContext("2d");
  canvas.height = viewport.height;
  canvas.width = viewport.width;

  await page.render({ canvasContext: ctx, viewport }).promise;
  el.pdfPageInfo.textContent = `${pageNum} / ${state.pdfDoc.numPages}`;
}

async function changePdfPage(delta) {
  if (!state.pdfDoc) return;
  const next = state.pdfPageNum + delta;
  if (next >= 1 && next <= state.pdfDoc.numPages) {
    state.pdfPageNum = next;
    await renderPdfPage(state.pdfPageNum);
  }
}

function openZoomModal() {
  const rec = getCurrentReceipt();
  if (!rec || !rec.dataUrl) return;
  el.zoomedImg.src = rec.dataUrl;
  el.zoomModal.classList.remove("hidden");
}

function renderItemsTable(items) {
  el.itemsTableBody.innerHTML = items
    .map(
      (item, idx) => `
      <tr>
        <td>
          <input
            type="text"
            class="table-input"
            value="${escapeHtml(item.item_name)}"
            oninput="updateLineItem(${idx}, 'item_name', this.value)"
          />
        </td>
        <td>
          <input
            type="number"
            step="any"
            class="table-input num-input"
            value="${item.qty}"
            oninput="updateLineItem(${idx}, 'qty', this.value)"
          />
        </td>
        <td>
          <input
            type="number"
            step="0.01"
            class="table-input num-input"
            value="${item.unit_price}"
            oninput="updateLineItem(${idx}, 'unit_price', this.value)"
          />
        </td>
        <td>
          <input
            type="number"
            step="0.01"
            class="table-input num-input"
            id="lineTotalInput_${idx}"
            value="${item.line_total}"
            oninput="updateLineItem(${idx}, 'line_total', this.value)"
          />
        </td>
        <td style="text-align:center;">
          <button
            type="button"
            class="btn btn-danger-ghost"
            onclick="deleteLineItem(${idx})"
            aria-label="Delete row"
          >
            🗑️
          </button>
        </td>
      </tr>
    `
    )
    .join("");
}

window.updateLineItem = function (idx, field, rawVal) {
  const rec = getCurrentReceipt();
  if (!rec || !rec.data.items[idx]) return;

  const item = rec.data.items[idx];
  if (field === "item_name") {
    item.item_name = rawVal;
  } else {
    item[field] = parseFloat(rawVal) || 0;
    // Automatically update line_total when qty or unit_price changes
    if (field === "qty" || field === "unit_price") {
      item.line_total = round2(item.qty * item.unit_price);
      const lineTotalEl = document.getElementById(`lineTotalInput_${idx}`);
      if (lineTotalEl) lineTotalEl.value = item.line_total.toFixed(2);
    }
  }

  runAuditAndRefreshUI();
};

window.deleteLineItem = function (idx) {
  const rec = getCurrentReceipt();
  if (!rec) return;
  rec.data.items.splice(idx, 1);
  renderItemsTable(rec.data.items);
  runAuditAndRefreshUI();
};

function recalculateTotals(mode) {
  const rec = getCurrentReceipt();
  if (!rec) return;

  const itemsSum = round2(
    rec.data.items.reduce((acc, item) => acc + round2(item.qty * item.unit_price), 0)
  );

  if (mode === "inclusive") {
    // Items already include 7% VAT
    const grand = itemsSum;
    const sub = round2(grand / 1.07);
    const vat = round2(grand - sub);
    rec.data.subtotal_ex_vat = sub;
    rec.data.vat_7_percent = vat;
    rec.data.grand_total = grand;
  } else {
    // Items exclude 7% VAT
    const sub = itemsSum;
    const vat = round2(sub * 0.07);
    const grand = round2(sub + vat);
    rec.data.subtotal_ex_vat = sub;
    rec.data.vat_7_percent = vat;
    rec.data.grand_total = grand;
  }

  el.fieldSubtotal.value = rec.data.subtotal_ex_vat.toFixed(2);
  el.fieldVat.value = rec.data.vat_7_percent.toFixed(2);
  el.fieldGrandTotal.value = rec.data.grand_total.toFixed(2);

  runAuditAndRefreshUI();
}

function runAuditAndRefreshUI() {
  const rec = getCurrentReceipt();
  if (!rec) return;

  rec.audit = validateReceiptMath(rec.data);
  renderValidationBanner(rec);
  renderReceiptSelector();
  updateDashboardMetrics();
}

function renderValidationBanner(rec) {
  const { isPass, status, issues, notes } = rec.audit;

  el.validationCard.className = `card validation-card ${isPass ? "pass" : "mismatch"}`;
  el.validationBadge.textContent = status;
  el.extractionTimeBadge.textContent = `⏱️ ${rec.extractionTimeSec.toFixed(2)}s`;

  if (isPass) {
    const noteItems = notes.map((n) => `<li>${escapeHtml(formatAuditMsg(n))}</li>`).join("");
    el.validationDetails.innerHTML = `
      <strong>${escapeHtml(t("auditPassSummary"))}</strong>
      ${noteItems ? `<ul>${noteItems}</ul>` : ""}
    `;
  } else {
    const issueItems = issues.map((iss) => `<li>${escapeHtml(formatAuditMsg(iss))}</li>`).join("");
    el.validationDetails.innerHTML = `<ul>${issueItems}</ul>`;
  }
}

// ============================================================================
// 10. System Performance Metrics & Ground Truth Evaluator
// ============================================================================
function updateDashboardMetrics() {
  const count = state.receipts.length;
  const totalSpend = round2(state.receipts.reduce((sum, r) => sum + r.data.grand_total, 0));
  const totalVat = round2(state.receipts.reduce((sum, r) => sum + r.data.vat_7_percent, 0));
  const passCount = state.receipts.filter((r) => r.audit.isPass).length;
  const passRate = count > 0 ? ((passCount / count) * 100).toFixed(1) : "0";

  el.kpiReceiptsCount.textContent = count;
  el.kpiNetSpend.textContent = `฿${formatCurrency(totalSpend)}`;
  el.kpiTotalVat.textContent = `฿${formatCurrency(totalVat)}`;
  el.kpiPassRate.textContent = `${passRate}%`;
  el.kpiPassCount.textContent = `${passCount} / ${count} PASS`;

  // Processing Time Comparison (AI vs 90s/receipt Manual Benchmark)
  const totalAiSec = round2(state.receipts.reduce((sum, r) => sum + r.extractionTimeSec, 0));
  const totalManualSec = count * 90;
  const savedSec = Math.max(0, round2(totalManualSec - totalAiSec));
  const savedPct =
    totalManualSec > 0 ? Math.max(0, ((savedSec / totalManualSec) * 100)).toFixed(1) : "0.0";

  el.metricAiTime.textContent = `${totalAiSec.toFixed(2)} s`;
  el.metricManualTime.textContent = `${totalManualSec.toFixed(2)} s`;

  const aiWidth = totalManualSec > 0 ? Math.min(100, (totalAiSec / totalManualSec) * 100) : 0;
  el.aiTimeBar.style.width = `${aiWidth}%`;
  el.manualTimeBar.style.width = count > 0 ? "100%" : "0%";

  el.metricTimeSavedPct.textContent = `${savedPct}%`;
  el.metricTimeSavedSec.textContent = `(-${savedSec.toFixed(1)}s)`;
}

function handleGroundTruthUpload(e) {
  const file = e.target.files[0];
  e.target.value = "";
  if (!file) return;

  if (state.receipts.length === 0) {
    showToast(t("toastGtNoData"), "error");
    return;
  }

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const workbook = XLSX.read(evt.target.result, { type: "array" });
      const summarySheet = workbook.Sheets["Receipt Summary"];
      if (!summarySheet) {
        showToast(t("toastGtError"), "error");
        return;
      }

      const gtSummaries = XLSX.utils.sheet_to_json(summarySheet);
      const detailsSheet = workbook.Sheets["Receipt Details"];
      const gtDetails = detailsSheet ? XLSX.utils.sheet_to_json(detailsSheet) : [];

      evaluateGroundTruth(gtSummaries, gtDetails);
    } catch (err) {
      console.error("Ground Truth parse error:", err);
      showToast(t("toastGtError"), "error");
    }
  };
  reader.readAsArrayBuffer(file);
}

function evaluateGroundTruth(gtSummaries, gtDetails) {
  let totalFieldsChecked = 0;
  let totalFieldsMatched = 0;
  const rows = [];

  state.receipts.forEach((rec, idx) => {
    // Match by Receipt No first, fallback to row index
    const gtRow =
      gtSummaries.find(
        (row) =>
          String(row["Receipt No"] || row["receipt_no"] || "").trim() ===
          rec.data.receipt_no.trim()
      ) || gtSummaries[idx];

    if (!gtRow) return;

    const fieldChecks = [
      {
        label: "Receipt No",
        actual: rec.data.receipt_no,
        expected: gtRow["Receipt No"] ?? gtRow["receipt_no"],
        isNum: false
      },
      {
        label: "Date",
        actual: rec.data.date,
        expected: gtRow["Date"] ?? gtRow["date"],
        isNum: false
      },
      {
        label: "Time",
        actual: rec.data.time,
        expected: gtRow["Time"] ?? gtRow["time"],
        isNum: false
      },
      {
        label: "Merchant",
        actual: rec.data.merchant_name,
        expected: gtRow["Merchant Name"] ?? gtRow["merchant_name"],
        isNum: false
      },
      {
        label: "Tax ID",
        actual: rec.data.tax_id,
        expected: gtRow["Tax ID"] ?? gtRow["tax_id"],
        isNum: false
      },
      {
        label: "Subtotal",
        actual: rec.data.subtotal_ex_vat,
        expected: gtRow["Subtotal (Excl. VAT)"] ?? gtRow["subtotal_ex_vat"],
        isNum: true
      },
      {
        label: "VAT 7%",
        actual: rec.data.vat_7_percent,
        expected: gtRow["VAT 7%"] ?? gtRow["vat_7_percent"],
        isNum: true
      },
      {
        label: "Grand Total",
        actual: rec.data.grand_total,
        expected: gtRow["Grand Total"] ?? gtRow["grand_total"],
        isNum: true
      }
    ];

    // Also check item count if Receipt Details sheet was supplied
    if (gtDetails.length > 0) {
      const expectedItems = gtDetails.filter(
        (d) =>
          String(d["Receipt No"] || d["receipt_no"] || "").trim() ===
          String(gtRow["Receipt No"] || gtRow["receipt_no"] || "").trim()
      );
      fieldChecks.push({
        label: "Item Count",
        actual: rec.data.items.length,
        expected: expectedItems.length,
        isNum: true
      });
    }

    let matched = 0;
    const diffs = [];

    fieldChecks.forEach((fc) => {
      totalFieldsChecked++;
      let isMatch = false;
      if (fc.isNum) {
        isMatch = Math.abs((Number(fc.actual) || 0) - (Number(fc.expected) || 0)) <= 0.05;
      } else {
        isMatch =
          String(fc.actual || "")
            .trim()
            .toLowerCase() ===
          String(fc.expected || "")
            .trim()
            .toLowerCase();
      }

      if (isMatch) {
        matched++;
        totalFieldsMatched++;
      } else {
        diffs.push(`${fc.label}: "${fc.actual}" ≠ "${fc.expected}"`);
      }
    });

    const accPct = ((matched / fieldChecks.length) * 100).toFixed(1);
    rows.push({
      receiptNo: rec.data.receipt_no || rec.fileName,
      matchedText: `${matched} / ${fieldChecks.length}`,
      accPct,
      diffsText: diffs.length > 0 ? diffs.join("; ") : "✅ 100% Exact Match"
    });
  });

  const overallAccuracy =
    totalFieldsChecked > 0 ? ((totalFieldsMatched / totalFieldsChecked) * 100).toFixed(2) : "0.00";

  state.groundTruthResults = { overallAccuracy, rows };
  renderGroundTruthTable();
}

function renderGroundTruthTable() {
  if (!state.groundTruthResults) return;
  el.gtResultBox.classList.remove("hidden");
  el.gtAccuracyValue.textContent = `${state.groundTruthResults.overallAccuracy}%`;

  el.gtTableBody.innerHTML = state.groundTruthResults.rows
    .map(
      (r) => `
      <tr>
        <td><strong>${escapeHtml(r.receiptNo)}</strong></td>
        <td>${escapeHtml(r.matchedText)}</td>
        <td><strong>${r.accPct}%</strong></td>
        <td>${escapeHtml(r.diffsText)}</td>
      </tr>
    `
    )
    .join("");
}

// ============================================================================
// 11. Export Deliverables (Excel .xlsx & UTF-8 JSON)
// ============================================================================
function exportToExcel() {
  if (state.receipts.length === 0) return;
  setWorkflowStage(5);

  // Sheet 1: Receipt Summary
  const summaryRows = state.receipts.map((rec) => ({
    "Receipt No": rec.data.receipt_no,
    Date: rec.data.date,
    Time: rec.data.time,
    "Merchant Name": rec.data.merchant_name,
    "Tax ID": rec.data.tax_id,
    "Item Count": rec.data.items.length,
    "Subtotal (Excl. VAT)": rec.data.subtotal_ex_vat,
    "VAT 7%": rec.data.vat_7_percent,
    "Grand Total": rec.data.grand_total,
    "Validation Status": rec.audit.isPass ? "PASS" : "MISMATCH",
    "AI Extraction Time (s)": rec.extractionTimeSec
  }));

  // Sheet 2: Receipt Details
  const detailRows = [];
  state.receipts.forEach((rec) => {
    rec.data.items.forEach((item, idx) => {
      detailRows.push({
        "Receipt No": rec.data.receipt_no,
        Date: rec.data.date,
        "Merchant Name": rec.data.merchant_name,
        "Line #": idx + 1,
        "Item Name": item.item_name,
        Quantity: item.qty,
        "Unit Price": item.unit_price,
        "Line Total": item.line_total
      });
    });
  });

  const wb = XLSX.utils.book_new();
  const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
  const wsDetails = XLSX.utils.json_to_sheet(detailRows);

  // Set readable column widths
  wsSummary["!cols"] = [
    { wch: 20 },
    { wch: 13 },
    { wch: 9 },
    { wch: 28 },
    { wch: 18 },
    { wch: 12 },
    { wch: 20 },
    { wch: 12 },
    { wch: 15 },
    { wch: 18 },
    { wch: 20 }
  ];

  wsDetails["!cols"] = [
    { wch: 20 },
    { wch: 13 },
    { wch: 28 },
    { wch: 8 },
    { wch: 34 },
    { wch: 12 },
    { wch: 14 },
    { wch: 14 }
  ];

  XLSX.utils.book_append_sheet(wb, wsSummary, "Receipt Summary");
  XLSX.utils.book_append_sheet(wb, wsDetails, "Receipt Details");

  const timestamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Accounting_Receipts_${timestamp}.xlsx`);
  showToast(t("toastExportExcel"), "success");
}

function exportToJson() {
  if (state.receipts.length === 0) return;
  setWorkflowStage(5);

  const payload = {
    exported_at: new Date().toISOString(),
    total_receipts: state.receipts.length,
    records: state.receipts.map((rec) => ({
      source_file: rec.fileName,
      extraction_time_seconds: rec.extractionTimeSec,
      validation_status: rec.audit.isPass ? "PASS" : "MISMATCH",
      ...rec.data
    }))
  };

  const jsonString = JSON.stringify(payload, null, 2);
  // Prepend UTF-8 BOM + explicit charset=utf-8 to preserve Thai characters cleanly
  const blob = new Blob(["\uFEFF" + jsonString], {
    type: "application/json;charset=utf-8"
  });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  const timestamp = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `Accounting_Receipts_${timestamp}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  showToast(t("toastExportJson"), "success");
}

// ============================================================================
// 12. Demo Sample Dataset & Utilities
// ============================================================================
function loadDemoDataset() {
  const demo1Data = {
    receipt_no: "TAX-20261124-0891",
    date: "24/11/2026",
    time: "14:28",
    merchant_name: "บริษัท ออฟฟิศซัพพลาย จำกัด (สำนักงานใหญ่)",
    tax_id: "0105561098421",
    items: [
      { item_name: "กระดาษถ่ายเอกสาร A4 80 แกรม (รีม)", qty: 5, unit_price: 120.0, line_total: 600.0 },
      { item_name: "หมึกพิมพ์เลเซอร์สีดำ TN-2380", qty: 1, unit_price: 1450.0, line_total: 1450.0 },
      { item_name: "แฟ้มสันกว้าง 3 นิ้ว สีดำ", qty: 4, unit_price: 62.5, line_total: 250.0 }
    ],
    subtotal_ex_vat: 2300.0,
    vat_7_percent: 161.0,
    grand_total: 2461.0
  };

  const demo2Data = {
    receipt_no: "INV-20261124-0412",
    date: "24/11/2026",
    time: "18:05",
    merchant_name: "ร้านกาแฟ อโรมา คอฟฟี่ เฮาส์",
    tax_id: "0105559044112",
    items: [
      { item_name: "Iced Americano (อเมริกาโน่เย็น)", qty: 2, unit_price: 75.0, line_total: 150.0 },
      { item_name: "Butter Croissant (ครัวซองต์เนยสด)", qty: 1, unit_price: 85.0, line_total: 85.0 }
    ],
    subtotal_ex_vat: 219.63,
    vat_7_percent: 15.37,
    grand_total: 245.0 // Intentional 10.00 THB discrepancy to showcase MISMATCH audit
  };

  state.receipts = [
    {
      id: "demo_1",
      fileName: "Demo_TaxInvoice_OfficeSupply.pdf",
      mimeType: "application/pdf",
      isDemo: true,
      demoHtml: `
        <div style="text-align:center;border-bottom:1px dashed #94a3b8;padding-bottom:8px;margin-bottom:8px;">
          <strong>ใบกำกับภาษี / ใบเสร็จรับเงิน</strong><br/>
          บริษัท ออฟฟิศซัพพลาย จำกัด (สำนักงานใหญ่)<br/>
          TAX ID: 0105561098421
        </div>
        <div>เลขที่: TAX-20261124-0891<br/>วันที่: 24/11/2026 14:28</div>
        <hr style="border:none;border-top:1px dashed #94a3b8;margin:8px 0;"/>
        <div>5x กระดาษ A4 80g @120.00 = 600.00</div>
        <div>1x หมึกเลเซอร์ TN-2380 = 1,450.00</div>
        <div>4x แฟ้มสันกว้าง 3" @62.50 = 250.00</div>
        <hr style="border:none;border-top:1px dashed #94a3b8;margin:8px 0;"/>
        <div style="text-align:right;">
          รวมก่อนภาษี: 2,300.00<br/>
          VAT 7%: 161.00<br/>
          <strong>ยอดสุทธิ: 2,461.00 บาท</strong>
        </div>
      `,
      extractionTimeSec: 1.38,
      data: demo1Data,
      audit: validateReceiptMath(demo1Data)
    },
    {
      id: "demo_2",
      fileName: "Demo_Cafe_Receipt_Mismatch.jpg",
      mimeType: "image/jpeg",
      isDemo: true,
      demoHtml: `
        <div style="text-align:center;border-bottom:1px dashed #94a3b8;padding-bottom:8px;margin-bottom:8px;">
          <strong>ใบกำกับภาษีอย่างย่อ (ABB)</strong><br/>
          ร้านกาแฟ อโรมา คอฟฟี่ เฮาส์<br/>
          TAX ID: 0105559044112
        </div>
        <div>เลขที่: INV-20261124-0412<br/>วันที่: 24/11/2026 18:05</div>
        <hr style="border:none;border-top:1px dashed #94a3b8;margin:8px 0;"/>
        <div>2x Iced Americano @75.00 = 150.00</div>
        <div>1x Butter Croissant @85.00 = 85.00</div>
        <hr style="border:none;border-top:1px dashed #94a3b8;margin:8px 0;"/>
        <div style="text-align:right;">
          Vatable: 219.63<br/>
          VAT 7%: 15.37<br/>
          <strong>TOTAL (OCR Misread): 245.00</strong>
        </div>
      `,
      extractionTimeSec: 1.14,
      data: demo2Data,
      audit: validateReceiptMath(demo2Data)
    }
  ];

  state.activeReceiptIndex = 0;
  setWorkflowStage(4);
  refreshFullWorkspace();
  showToast(t("toastDemoLoaded"), "success");
}

function clearAllData() {
  state.pendingFiles = [];
  state.receipts = [];
  state.activeReceiptIndex = 0;
  state.groundTruthResults = null;
  el.gtResultBox.classList.add("hidden");
  renderPendingFiles();
  setWorkflowStage(1);
  refreshFullWorkspace();
}

function formatCurrency(val) {
  return Number(val || 0).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function showToast(message, type = "info") {
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;
  el.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.remove();
  }, 4200);
}

// Start Application
document.addEventListener("DOMContentLoaded", init);
