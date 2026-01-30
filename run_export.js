const ExcelJS = require("exceljs");
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

const URL = "https://eprplastic.cpcb.gov.in/#/epr/details/sales";
function getConfigPath() {
    const idx = process.argv.indexOf("--config");
    if (idx !== -1 && process.argv[idx + 1]) {
        return path.resolve(__dirname, process.argv[idx + 1]);
    }
    return path.resolve(__dirname, "config.json");
}

const CONFIG_PATH = getConfigPath();

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
    const plasticType = String(cfg?.plasticType || "PP").trim();
    const storageState = String(cfg?.storageState || "storageState.json").trim();
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
    return { inputExcel, sheetName, outputExcel, maxRows, plasticType, storageState };
}

const CONFIG = loadConfig();
const STORAGE = path.resolve(__dirname, CONFIG.storageState);
const EXCEL_PATH = path.resolve(__dirname, CONFIG.inputExcel);
const SHEET = CONFIG.sheetName;
const OUTPUT_PATH = path.resolve(__dirname, CONFIG.outputExcel);
const EXCEL_TMP = `${OUTPUT_PATH}.tmp`;
const EXCEL_BAK = `${OUTPUT_PATH}.bak`;
const OUTPUT_BASENAME = path.basename(OUTPUT_PATH, path.extname(OUTPUT_PATH));
const LOG_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_log.csv`);

// ---------- Helpers ----------
function normHeader(s) {
    return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object" && v.text) return String(v.text).trim();
    return String(v).trim();
}

function csvEscape(v) {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) {
        return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
}

function ensureLogHeader() {
    if (fs.existsSync(LOG_PATH)) return;
    const header = [
        "datetime",
        "row",
        "e_invoice_number",
        "sales_date",
        "quantity_sold_mt",
        "seller_gst",
        "epr_invoice_number",
        "status",
        "message",
    ].join(",");
    fs.writeFileSync(LOG_PATH, `${header}\n`);
}

function appendLogRow(row, headerMap, { status, eprInvoiceNumber, message }) {
    ensureLogHeader();
    const ts = new Date().toISOString();
    const data = [
        ts,
        row.number,
        cellText(getVal(row, headerMap, "E-Invoice Number*")),
        cellText(getVal(row, headerMap, "Sales date*")),
        cellText(getVal(row, headerMap, "Quantity Sold(MT)")),
        cellText(getVal(row, headerMap, "GST No. of Seller *")),
        cellText(eprInvoiceNumber),
        status,
        message || "",
    ].map(csvEscape);
    fs.appendFileSync(LOG_PATH, `${data.join(",")}\n`);
}

function excelDateToISO(v) {
    if (!v) throw new Error("Sales date is empty");
    if (v instanceof Date) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, "0");
        const d = String(v.getDate()).padStart(2, "0");
        return `${y}-${m}-${d}`;
    }
    const s = cellText(v);
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const parts = s.split("/");
    if (parts.length === 3) {
        const mm = String(Number(parts[0])).padStart(2, "0");
        const dd = String(Number(parts[1])).padStart(2, "0");
        const yyyy = String(Number(parts[2]));
        return `${yyyy}-${mm}-${dd}`;
    }
    throw new Error(`Unsupported Sales date format: ${s}`);
}

function formatQty(v) {
    if (typeof v === "number") return v.toFixed(2);
    const s = cellText(v);
    const n = Number(s);
    if (Number.isFinite(n)) return n.toFixed(2);
    return s;
}

function randDelayMs(minMs = 3000, maxMs = 7000) {
    const min = Math.floor(minMs);
    const max = Math.floor(maxMs);
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

// ---------- Excel helpers ----------
function getHeaderMap(ws) {
    const headerRow = ws.getRow(1);
    const map = new Map();
    headerRow.eachCell((cell, colNumber) => {
        const key = normHeader(cellText(cell.value));
        if (key) map.set(key, colNumber);
    });
    return map;
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

async function syncInputWorkbook(wb) {
    if (path.resolve(EXCEL_PATH) === path.resolve(OUTPUT_PATH)) return;
    await safeWriteWorkbookToPath(wb, EXCEL_PATH);
}

// ---------- ERP helpers ----------
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

async function clickAddNew(page) {
    const addNewBtn = page.getByRole("button", { name: "Add New Export", exact: true }).first();
    await addNewBtn.waitFor({ state: "visible", timeout: 60000 });
    await addNewBtn.scrollIntoViewIfNeeded();
    await addNewBtn.click();
    await page.waitForTimeout(300);
}

async function selectCat2Row(page, plasticTypeText) {
    console.log(`selectCat2Row: start (plastic=${plasticTypeText || "ANY"})`);
    await page.waitForSelector("#ScrollableSimpleTableBody", { timeout: 60000 });
    let catRow = page.locator("tbody#ScrollableSimpleTableBody tr", {
        has: page.locator('span[title="CAT-II"]'),
    });

    if (plasticTypeText) {
        catRow = catRow.filter({
            has: page.locator(`span[title="${plasticTypeText}"]`),
        });
    }

    catRow = catRow.first();
    const rowCount = await catRow.count();
    console.log(`selectCat2Row: rowCount=${rowCount}`);
    const tableRows = await page.locator("tbody#ScrollableSimpleTableBody tr").count();
    console.log(`selectCat2Row: totalRows=${tableRows}`);
    const catCount = await page.locator('span[title="CAT-II"]').count();
    console.log(`selectCat2Row: CAT-II spans=${catCount}`);
    if (plasticTypeText) {
        const plasticCount = await page.locator(`span[title="${plasticTypeText}"]`).count();
        console.log(`selectCat2Row: ${plasticTypeText} spans=${plasticCount}`);
    }
    await catRow.waitFor({ state: "visible", timeout: 20000 });
    const checkbox = catRow.locator('input[type="checkbox"][name="check-box"]').first();
    await checkbox.scrollIntoViewIfNeeded();
    await checkbox.click({ force: true });
    await page.waitForSelector('input[name="qty_product_sold"]', { timeout: 30000 });
    console.log("selectCat2Row: checkbox clicked, form visible");
}

async function setAngularDateById(page, id, isoDate) {
    await page.evaluate(
        ({ id, isoDate }) => {
            const el = document.getElementById(id);
            if (!el) throw new Error(`Element #${id} not found`);
            el.value = isoDate;
            el.dispatchEvent(new Event("input", { bubbles: true }));
            el.dispatchEvent(new Event("change", { bubbles: true }));
            el.blur();
        },
        { id, isoDate }
    );
}

async function fillBySelectorIfEnabled(page, selector, value) {
    const v = cellText(value);
    const loc = page.locator(selector).first();
    await loc.waitFor({ state: "visible", timeout: 60000 });
    await loc.scrollIntoViewIfNeeded();
    const disabled = await loc.evaluate((el) => el.disabled || el.readOnly).catch(() => false);
    if (disabled) return;
    await loc.click({ timeout: 15000 });
    await loc.fill("");
    if (v) await loc.fill(v);
    await loc.blur();
}

async function fillByIdIfEnabled(page, id, value) {
    const v = cellText(value);
    const loc = page.locator(`#${id}`);
    await loc.waitFor({ state: "visible", timeout: 60000 });
    await loc.scrollIntoViewIfNeeded();
    const disabled = await loc.evaluate((el) => el.disabled || el.readOnly).catch(() => false);
    if (disabled) return;
    await loc.click({ timeout: 15000 });
    await loc.fill("");
    if (v) await loc.fill(v);
    await loc.blur();
}

async function clickSubmitAndConfirm(page) {
    const submit = page.locator('button[type="submit"]', { hasText: "Generate EPR Invoice Number" }).first();
    await submit.waitFor({ state: "visible", timeout: 20000 });
    if (await submit.isDisabled()) {
        throw new Error("Submit disabled: some required fields still missing.");
    }
    await submit.click();
    try {
        const confirmBtn = page.locator(".modal-footer button", { hasText: "Confirm" }).first();
        await confirmBtn.waitFor({ state: "visible", timeout: 60000 });
        await confirmBtn.click();
    } catch { }
}

async function clickResetAndConfirm(page) {
    const reset = page.locator("button", { hasText: /\bReset\b/i }).first();
    if (!(await reset.count())) return false;
    await reset.waitFor({ state: "visible", timeout: 10000 }).catch(() => { });
    await reset.scrollIntoViewIfNeeded().catch(() => { });
    await reset.click().catch(() => { });

    const modal = page.locator(".modal-dialog, .modal-content").first();
    if (await modal.count()) {
        try {
            await modal.waitFor({ state: "visible", timeout: 8000 });
            const confirmBtn = modal.getByRole("button", { name: "Confirm", exact: true }).first();
            if (await confirmBtn.count()) {
                await confirmBtn.click();
            }
        } catch { }
    }
    return true;
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

async function readEprInvoiceNumber(page) {
    const input = page.locator("#invoiceNumberCopy").first();
    if (await input.count()) {
        try {
            const val = (await input.inputValue()).trim();
            if (val) return val;
        } catch { }
    }

    const label = page.locator("text=/EPR\\s*Invoice\\s*Number/i").first();
    if (await label.count()) {
        const container = label.locator("xpath=ancestor-or-self::*[self::div or self::span or self::p][1]");
        const text = (await container.innerText().catch(() => "")) || "";
        const match = text.match(/EPR\s*Invoice\s*Number\s*[:\\-]?\\s*([A-Za-z0-9\\-\\/]+)/i);
        if (match && match[1]) return match[1].trim();
        const sibling = label.locator("xpath=following::span[1] | following::div[1] | following::p[1]").first();
        const sibText = (await sibling.innerText().catch(() => "")).trim();
        if (sibText) return sibText;
    }
    return "";
}

// ---------- Main ----------
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
    await clickAddNew(page);

    const lastRow = CONFIG.maxRows ? Math.min(ws.rowCount, CONFIG.maxRows) : ws.rowCount;
    for (let r = 2; r <= lastRow; r++) {
        const row = ws.getRow(r);
        if (isRowEmpty(row, headerMap)) {
            console.log(`Row ${r}: Skipped (row empty)`);
            continue;
        }

        const status = cellText(getVal(row, headerMap, "Status"));
        if (status.toLowerCase().includes("success") || status.toLowerCase().includes("filled")) {
            continue;
        }

        const eprInvoiceExisting = getVal(row, headerMap, "EPR Invoice Number");
        if (!isCellEmpty(eprInvoiceExisting)) {
            console.log(`Row ${r}: Skipped (EPR Invoice already present)`);
            continue;
        }

        const qtySold = getVal(row, headerMap, "Quantity Sold(MT)");
        const entityName = getVal(row, headerMap, "Name of the Entity *");
        const entityAddress = getVal(row, headerMap, "Address*");
        const sellerGst = getVal(row, headerMap, "GST No. of Seller *");
        const hsn = getVal(row, headerMap, "HSN CODE");
        const invno = getVal(row, headerMap, "E-Invoice Number*");
        const account = getVal(row, headerMap, "Bank Account No*");
        const ifsc = getVal(row, headerMap, "IFSC Code*");
        const principal = getVal(row, headerMap, "Principal Amount(₹)*");
        const gstOther = getVal(row, headerMap, "GST & Other Charges(₹)*");
        const salesDateRaw = getVal(row, headerMap, "Sales date*");

        try {
            console.log(`Row ${r} starting...`);

            await selectCat2Row(page, CONFIG.plasticType || "PP");

            await fillBySelectorIfEnabled(page, 'input[name="qty_product_sold"]', formatQty(qtySold));
            await fillBySelectorIfEnabled(page, 'input[formcontrolname="entity_name"]', entityName);
            await fillBySelectorIfEnabled(page, 'input[formcontrolname="entity_address"]', entityAddress);
            await fillByIdIfEnabled(page, "sellerGst", sellerGst);
            await fillByIdIfEnabled(page, "hsnCode", hsn);
            await fillByIdIfEnabled(page, "invno", invno);
            await fillByIdIfEnabled(page, "account_number", account);
            await fillByIdIfEnabled(page, "ifsc_code", ifsc);
            await fillByIdIfEnabled(page, "amount", principal);
            await fillByIdIfEnabled(page, "gst", gstOther);

            const salesDateISO = excelDateToISO(salesDateRaw);
            await setAngularDateById(page, "salesDate", salesDateISO);

            await clickSubmitAndConfirm(page);
            await page.waitForTimeout(1000);
            await waitForLoaderToFinish(page);
            const toastText = await readToastText(page);

            const eprInvoice = await readEprInvoiceNumber(page);
            if (!eprInvoice) {
                throw new Error("EPR Invoice Number not found after submit.");
            }

            setVal(row, headerMap, "Status", "Filled");
            setVal(row, headerMap, "EPR Invoice Number", eprInvoice);
            row.commit();
            await safeWriteWorkbook(wb);
            await syncInputWorkbook(wb);

            appendLogRow(row, headerMap, {
                status: "Filled",
                eprInvoiceNumber: eprInvoice,
                message: toastText,
            });

            console.log(`Row ${r}: Filled ✓`);
            const delayMs = randDelayMs(3000, 7000);
            const startTs = new Date().toISOString();
            console.log(`Row ${r}: delay start ${startTs} (${delayMs}ms)`);
            await page.waitForTimeout(delayMs);
            const endTs = new Date().toISOString();
            console.log(`Row ${r}: delay end ${endTs}`);
        } catch (e) {
            const msg = String(e?.message || e);
            console.log(`Row ${r}: Failed ❌ ->`, msg);

            setVal(row, headerMap, "Status", "Failed: " + msg);
            row.commit();
            await safeWriteWorkbook(wb);
            await syncInputWorkbook(wb);

            appendLogRow(row, headerMap, {
                status: "Failed",
                eprInvoiceNumber: "",
                message: msg,
            });
        } finally {
            if (page.isClosed()) {
                console.log("Page closed. Stopping.");
                break;
            }
            await waitForLoaderToFinish(page);
            const didReset = await clickResetAndConfirm(page);
            if (!didReset) {
                await clickAddNew(page);
            }
        }
    }

    await browser.close();
    console.log("Done. Updated Excel:", EXCEL_PATH);
})();
