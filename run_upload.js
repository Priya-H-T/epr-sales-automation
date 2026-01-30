const ExcelJS = require("exceljs");
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

const URL = "https://eprplastic.cpcb.gov.in/#/epr/details/sales";
const CONFIG_PATH = path.resolve(__dirname, "config.json");
const STORAGE = path.resolve(__dirname, "storageState.json");
const ROOT_DIR = __dirname;

function loadConfig() {
    if (!fs.existsSync(CONFIG_PATH)) {
        throw new Error(`Missing config file: ${CONFIG_PATH}`);
    }
    const raw = fs.readFileSync(CONFIG_PATH, "utf8");
    const cfg = JSON.parse(raw);
    const inputExcel = String(cfg?.inputExcel || "").trim();
    const sheetName = String(cfg?.sheetName || "").trim();
    const outputExcel = String(cfg?.outputExcel || "").trim();
    const maxRowsRaw = cfg?.max_rows;
    const invoicePdfDir = String(cfg?.invoicePdfDir || "").trim();
    if (!inputExcel || !sheetName || !outputExcel) {
        throw new Error("config.json must include inputExcel, sheetName, and outputExcel");
    }
    let maxRows = null;
    if (maxRowsRaw !== undefined && maxRowsRaw !== null && String(maxRowsRaw).trim() !== "") {
        const n = Number(maxRowsRaw);
        if (!Number.isFinite(n)) {
            throw new Error("config.json max_rows must be a number when provided");
        }
        if (n > 0) {
            maxRows = Math.floor(n);
        }
    }
    return { inputExcel, sheetName, outputExcel, maxRows, invoicePdfDir };
}

const CONFIG = loadConfig();
const EXCEL_PATH = path.resolve(__dirname, CONFIG.inputExcel);
const SHEET = CONFIG.sheetName;
const OUTPUT_PATH = path.resolve(__dirname, CONFIG.outputExcel);
const EXCEL_TMP = `${OUTPUT_PATH}.tmp`;
const EXCEL_BAK = `${OUTPUT_PATH}.bak`;
const OUTPUT_BASENAME = path.basename(OUTPUT_PATH, path.extname(OUTPUT_PATH));
const UPLOAD_LOG_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_upload_log.csv`);
const UPLOAD_FILLED_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_upload_filled.csv`);

function normHeader(s) {
    return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object" && v.text) return String(v.text).trim();
    return String(v).trim();
}

function getHeaderMap(ws) {
    const headerRow = ws.getRow(1);
    const map = new Map();
    headerRow.eachCell((cell, colNumber) => {
        const key = normHeader(cellText(cell.value));
        if (key) map.set(key, colNumber);
    });
    return map;
}

function getHeaderList(ws) {
    const headerRow = ws.getRow(1);
    const headers = [];
    headerRow.eachCell((cell) => {
        const key = cellText(cell.value);
        if (key) headers.push(key);
    });
    return headers;
}

function getVal(row, headerMap, headerName) {
    const col = headerMap.get(normHeader(headerName));
    if (!col) return "";
    return row.getCell(col).value;
}

function setVal(row, headerMap, headerName, value) {
    const col = headerMap.get(normHeader(headerName));
    if (!col) return;
    row.getCell(col).value = value;
}

function isCellEmpty(v) {
    return cellText(v) === "";
}

function isRowEmpty(row, headerMap) {
    for (const col of headerMap.values()) {
        if (!isCellEmpty(row.getCell(col).value)) return false;
    }
    return true;
}

function csvEscape(v) {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) {
        return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
}

function ensureUploadLogHeader() {
    if (fs.existsSync(UPLOAD_LOG_PATH)) return;
    const header = [
        "datetime",
        "row",
        "e_invoice_number",
        "epr_invoice_number",
        "status",
        "message",
    ].join(",");
    fs.writeFileSync(UPLOAD_LOG_PATH, `${header}\n`);
}

function appendUploadLogRow(row, headerMap, { status, message }) {
    ensureUploadLogHeader();
    const ts = new Date().toISOString();
    const data = [
        ts,
        row.number,
        cellText(getVal(row, headerMap, "E-Invoice Number*")),
        cellText(getVal(row, headerMap, "EPR Invoice Number")),
        status,
        message || "",
    ].map(csvEscape);
    fs.appendFileSync(UPLOAD_LOG_PATH, `${data.join(",")}\n`);
}

function ensureUploadFilledHeader(headerList) {
    if (fs.existsSync(UPLOAD_FILLED_PATH)) return;
    const header = [...headerList, "datetime", "message"].map(csvEscape).join(",");
    fs.writeFileSync(UPLOAD_FILLED_PATH, `${header}\n`);
}

function appendUploadFilledRow(row, headerMap, headerList, { message }) {
    ensureUploadFilledHeader(headerList);
    const ts = new Date().toISOString();
    const rowData = headerList.map((h) => cellText(getVal(row, headerMap, h)));
    const data = [...rowData, ts, message || ""].map(csvEscape);
    fs.appendFileSync(UPLOAD_FILLED_PATH, `${data.join(",")}\n`);
}

async function safeWriteWorkbook(wb) {
    await wb.xlsx.writeFile(EXCEL_TMP);
    try {
        if (fs.existsSync(OUTPUT_PATH) && fs.statSync(OUTPUT_PATH).size > 0) {
            fs.copyFileSync(OUTPUT_PATH, EXCEL_BAK);
        }
    } catch { }
    fs.renameSync(EXCEL_TMP, OUTPUT_PATH);
}

async function safeWriteWorkbookToPath(wb, targetPath) {
    const tmp = `${targetPath}.tmp`;
    const bak = `${targetPath}.bak`;
    await wb.xlsx.writeFile(tmp);
    try {
        if (fs.existsSync(targetPath) && fs.statSync(targetPath).size > 0) {
            fs.copyFileSync(targetPath, bak);
        }
    } catch { }
    fs.renameSync(tmp, targetPath);
}

async function appendOutputToInput({ inputPath, outputPath, sheetName }) {
    const inPath = path.resolve(inputPath);
    const outPath = path.resolve(outputPath);
    if (inPath === outPath) {
        console.log("Append skipped: input and output are the same file.");
        return;
    }
    if (!fs.existsSync(outPath)) {
        console.log(`Append skipped: output file not found: ${outPath}`);
        return;
    }
    const wbIn = new ExcelJS.Workbook();
    const wbOut = new ExcelJS.Workbook();
    await wbIn.xlsx.readFile(inPath);
    await wbOut.xlsx.readFile(outPath);
    const wsIn = wbIn.getWorksheet(sheetName);
    const wsOut = wbOut.getWorksheet(sheetName);
    if (!wsIn || !wsOut) {
        throw new Error(`Append failed: sheet not found (${sheetName})`);
    }
    for (let r = 2; r <= wsOut.rowCount; r++) {
        const rowOut = wsOut.getRow(r);
        if (!rowOut.hasValues) continue;
        const newRow = wsIn.addRow(rowOut.values);
        newRow.commit();
    }
    await safeWriteWorkbookToPath(wbIn, inPath);
    console.log("Append complete: output appended to input.");
}

async function waitForLoaderToFinish(page) {
    const loaders = [
        ".spinner-border",
        ".loading",
        ".loader",
        ".ngx-spinner-overlay",
        ".ngx-spinner",
        ".overlay",
        ".block-ui-wrapper",
        ".k-i-loading",
    ];
    for (const sel of loaders) {
        try {
            const loc = page.locator(sel);
            if ((await loc.count()) > 0) {
                await loc.first().waitFor({ state: "hidden", timeout: 30000 }).catch(() => { });
            }
        } catch { }
    }
}

async function readToastText(page) {
    const toast = page.locator(".toast, .toaster, .ngx-toastr, .toast-container").first();
    if (!(await toast.count())) return "";
    try {
        const text = (await toast.innerText()).trim();
        return text;
    } catch {
        return "";
    }
}

async function readToastTextWithRetry(page, retries = 6, delayMs = 300) {
    for (let i = 0; i < retries; i++) {
        const text = await readToastText(page);
        if (text) return text;
        await page.waitForTimeout(delayMs);
    }
    return "";
}

function isSuccessText(text) {
    const t = String(text || "");
    return /success/i.test(t) && !/error/i.test(t);
}

function normalizeStatus(text) {
    return String(text || "").trim().toLowerCase();
}

function findPdfByInvoiceNumber(rootDir, invoiceNumber) {
    const target = String(invoiceNumber || "").trim();
    if (!target) return "";
    const targetLower = target.toLowerCase();
    const skip = new Set(["node_modules", ".git", ".vscode"]);

    const stack = [rootDir];
    const matches = [];

    while (stack.length) {
        const dir = stack.pop();
        let entries = [];
        try {
            entries = fs.readdirSync(dir, { withFileTypes: true });
        } catch {
            continue;
        }
        for (const ent of entries) {
            const full = path.join(dir, ent.name);
            if (ent.isDirectory()) {
                if (!skip.has(ent.name)) stack.push(full);
                continue;
            }
            if (!ent.isFile()) continue;
            const nameLower = ent.name.toLowerCase();
            if (nameLower.endsWith(".pdf") && nameLower.includes(targetLower)) {
                matches.push(full);
            }
        }
    }

    if (matches.length === 0) return "";
    matches.sort((a, b) => a.length - b.length);
    return matches[0];
}

async function getTableColumnIndex(page, headerText) {
    const idx = await page.evaluate((headerText) => {
        const norm = (s) => (s || "").replace(/\s+/g, " ").trim().toLowerCase();
        const ths = Array.from(document.querySelectorAll("#simple_table_header th"));
        const target = norm(headerText);
        const i = ths.findIndex((th) => norm(th.innerText) === target);
        return i >= 0 ? i + 1 : null;
    }, headerText);
    return idx;
}

async function searchSalesByEprInvoice(page, eprInvoiceNumber) {
    console.log(`[search] start for EPR ${eprInvoiceNumber}`);
    // Ensure any modal/overlay is gone before searching
    await page.locator(".modal-dialog").first().waitFor({ state: "hidden", timeout: 2000 }).catch(() => { });
    await waitForLoaderToFinish(page);
    const searchInput = page.locator('input[name="searchField"]').first();
    await searchInput.waitFor({ state: "visible", timeout: 30000 });
    await page
        .waitForFunction((el) => !el.disabled && !el.readOnly, await searchInput.elementHandle(), {
            timeout: 5000,
        })
        .catch(() => { });
    console.log("[search] input visible");
    await searchInput.click();
    await searchInput.fill("");
    await searchInput.fill(eprInvoiceNumber);

    const searchBtn = page.locator("button", { hasText: "Search" }).first();
    await searchBtn.click();
    console.log("[search] clicked");
    await page.waitForTimeout(1500);
    await waitForLoaderToFinish(page);

    const row = page.locator("#ScrollableSimpleTableBody tr", { hasText: eprInvoiceNumber }).first();
    await row.waitFor({ state: "visible", timeout: 20000 });
    console.log("[search] row visible");
    return row;
}

async function openUploadModalForRow(page, { eprInvoiceNumber, eInvoiceNumber, rootDir }) {
    if (!eprInvoiceNumber) {
        throw new Error("Missing EPR Invoice Number for upload.");
    }
    if (!eInvoiceNumber) {
        throw new Error("Missing E-Invoice Number for upload.");
    }

    const row = await searchSalesByEprInvoice(page, eprInvoiceNumber);
    const colIdx = await getTableColumnIndex(page, "Invoice File Status");
    if (!colIdx) {
        throw new Error("Invoice File Status column not found.");
    }

    const statusCell = row.locator(`td:nth-child(${colIdx})`).first();
    const redIcon = statusCell.locator(".fa-exclamation-triangle.color-red").first();
    const greenIcon = statusCell.locator(".color-green, .fa-check, .fa-check-circle").first();

    if (await greenIcon.count()) {
        return { status: "already", toast: "Already uploaded (green status)." };
    }
    if (!(await redIcon.count())) {
        return { status: "skipped", toast: "Invoice File Status not red." };
    }

    await redIcon.click();

    const modal = page.locator('.modal-dialog:has-text("Upload Invoice")').first();
    await modal.waitFor({ state: "visible", timeout: 20000 });

    const filePath = findPdfByInvoiceNumber(rootDir, eInvoiceNumber);
    if (!filePath) {
        throw new Error(`Invoice PDF not found for E-Invoice Number: ${eInvoiceNumber}`);
    }
    if (!fs.existsSync(filePath)) {
        throw new Error(`Invoice PDF path does not exist: ${filePath}`);
    }

    const fileInput = modal.locator('input[type="file"][name="invoice"]').first();
    await fileInput.setInputFiles(filePath);
    await page.waitForFunction((el) => el && el.files && el.files.length > 0, await fileInput.elementHandle(), {
        timeout: 15000,
    });

    return { status: "ready", modal, redIcon };
}

async function uploadInvoiceFileWithRetry(page, opts) {
    const prep = await openUploadModalForRow(page, opts);
    if (prep.status === "already" || prep.status === "skipped") {
        return { status: prep.status, toast: prep.toast, attempts: 0 };
    }

    let modal = prep.modal;
    let uploadBtn = modal.locator("button.btn.btn-primary", { hasText: "Upload" }).first();
    const ensureUploadReady = async () => {
        console.log("ensureUploadReady: start");
        if (!(await modal.isVisible().catch(() => false))) {
            console.log("ensureUploadReady: modal not visible, reopening");
            await prep.redIcon.click();
            modal = page.locator('.modal-dialog:has-text("Upload Invoice")').first();
            await modal.waitFor({ state: "visible", timeout: 5000 }).catch(() => { });
            uploadBtn = modal.locator("button.btn.btn-primary", { hasText: "Upload" }).first();
        }
        await waitForLoaderToFinish(page);
        await page.locator("#loader-wrapper").first().waitFor({ state: "hidden", timeout: 2000 }).catch(() => { });
        const uploadHandle = await uploadBtn.elementHandle();
        if (uploadHandle) {
            await page.waitForFunction((el) => !el.disabled, uploadHandle, { timeout: 2000 }).catch(() => { });
        }
        console.log("ensureUploadReady: upload enabled");
    };

    let lastToast = "";
    let lastStatus = "error";
    for (let attempt = 1; attempt <= 2; attempt++) {
        console.log(`Upload attempt ${attempt} for ${opts.eprInvoiceNumber}`);
        await ensureUploadReady();
        try {
            await uploadBtn.click();
        } catch {
            // Loader may intercept; try force click once
            await uploadBtn.click({ force: true });
        }
        console.log("Upload clicked");
        await page.waitForTimeout(1200);
        const toastText = await readToastTextWithRetry(page, 6, 300);
        console.log(`Toast read: ${toastText ? "yes" : "no"}`);
        lastToast = toastText;
        lastStatus = isSuccessText(toastText) ? "success" : "error";
        if (lastStatus === "success") break;
    }

    // Close modal after attempts
    try {
        const closeBtn = modal.locator("button", { hasText: "Close" }).first();
        if (await closeBtn.count()) {
            await closeBtn.click();
            console.log("Modal close button clicked");
        } else {
            const closeIcon = modal.locator("#closeInvoiceUploadPopup, .close").first();
            if (await closeIcon.count()) {
                await closeIcon.click();
                console.log("Modal close icon clicked");
            }
        }
        // Keep this short so success doesn't stall
        await modal.waitFor({ state: "hidden", timeout: 3000 }).catch(() => { });
        console.log("Modal hidden");
    } catch { }

    return { status: lastStatus, toast: lastToast, attempts: 2 };
}

(async () => {
    if (!fs.existsSync(EXCEL_PATH)) {
        throw new Error(`Excel not found: ${EXCEL_PATH}`);
    }
    const stat = fs.statSync(EXCEL_PATH);
    if (stat.size < 1000) {
        throw new Error(`Excel looks empty/corrupt (${stat.size} bytes): ${EXCEL_PATH}`);
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(EXCEL_PATH);

    const ws = wb.getWorksheet(SHEET);
    if (!ws) {
        throw new Error(`Sheet not found: ${SHEET}`);
    }

    const headerMap = getHeaderMap(ws);
    const headerList = getHeaderList(ws);

    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext(
        fs.existsSync(STORAGE) ? { storageState: STORAGE } : {}
    );
    const page = await context.newPage();

    await page.goto(URL, { waitUntil: "domcontentloaded" });

    if (!fs.existsSync(STORAGE)) {
        console.log("Login manually in this Playwright window, then press ENTER here...");
        await new Promise((res) => process.stdin.once("data", () => res()));
        await context.storageState({ path: STORAGE });
        console.log("Saved session to storageState.json");
    }

    await page.goto(URL, { waitUntil: "domcontentloaded" });
    await page.waitForSelector("#ScrollableSimpleTableBody", { timeout: 60000 });

    const lastRow = CONFIG.maxRows ? Math.min(ws.rowCount, CONFIG.maxRows) : ws.rowCount;
    const pdfRoot = CONFIG.invoicePdfDir
        ? path.resolve(__dirname, CONFIG.invoicePdfDir)
        : ROOT_DIR;

    for (let r = 2; r <= lastRow; r++) {
        const row = ws.getRow(r);
        if (isRowEmpty(row, headerMap)) {
            console.log(`Row ${r}: Skipped (row empty)`);
            continue;
        }

        const eprInvoice = cellText(getVal(row, headerMap, "EPR Invoice Number"));
        const eInvoice = cellText(getVal(row, headerMap, "E-Invoice Number*"));
        const uploadStatusRaw = cellText(getVal(row, headerMap, "Invoice update Status"));
        const uploadStatus = normalizeStatus(uploadStatusRaw);

        if (!eprInvoice) {
            console.log(`Row ${r}: Skipped (missing EPR Invoice Number)`);
            continue;
        }
        if (uploadStatusRaw.trim() !== "" || uploadStatus.includes("success") || uploadStatus.includes("sucess")) {
            console.log(`Row ${r}: Invoice update Status already filled, skipping upload`);
            continue;
        }

        try {
            const uploadResult = await uploadInvoiceFileWithRetry(page, {
                eprInvoiceNumber: eprInvoice,
                eInvoiceNumber: eInvoice,
                rootDir: pdfRoot,
            });

            const uploadMessage = uploadResult.toast || uploadResult.status;
            if (uploadResult.status === "already") {
                setVal(row, headerMap, "Invoice update Status", "Sucess");
                appendUploadLogRow(row, headerMap, { status: "Success", message: uploadMessage });
                appendUploadFilledRow(row, headerMap, headerList, { message: uploadMessage });
            } else if (uploadResult.status === "success") {
                setVal(row, headerMap, "Invoice update Status", "Sucess");
                appendUploadLogRow(row, headerMap, { status: "Success", message: uploadMessage });
                appendUploadFilledRow(row, headerMap, headerList, { message: uploadMessage });
            } else {
                setVal(row, headerMap, "Invoice update Status", "Failed: " + (uploadMessage || "Error"));
                appendUploadLogRow(row, headerMap, { status: "Failed", message: uploadMessage });
                appendUploadFilledRow(row, headerMap, headerList, { message: uploadMessage });
            }
        } catch (e) {
            const msg = String(e?.message || e);
            setVal(row, headerMap, "Invoice update Status", "Failed: " + msg);
            appendUploadLogRow(row, headerMap, { status: "Failed", message: msg });
            appendUploadFilledRow(row, headerMap, headerList, { message: msg });
        }

        row.commit();
        await safeWriteWorkbook(wb);
    }

    await browser.close();
    try {
        await appendOutputToInput({
            inputPath: EXCEL_PATH,
            outputPath: OUTPUT_PATH,
            sheetName: SHEET,
        });
    } catch (e) {
        console.log("Append output to input failed:", String(e?.message || e));
    }
    console.log("Done. Updated Excel:", EXCEL_PATH);
})();
