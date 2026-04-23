const ExcelJS = require("exceljs");
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

const PIBO_URL = "https://eprplastic.cpcb.gov.in/#/epr/pibo-operations/sales";

function getConfigPath() {
    const idx = process.argv.indexOf("--config");
    if (idx !== -1 && process.argv[idx + 1]) {
        return path.resolve(__dirname, process.argv[idx + 1]);
    }
    return path.resolve(__dirname, "config_pibo.json");
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
    const storageState = String(cfg?.storageState || "storageState_pibo.json").trim();
    if (!inputExcel || !sheetName || !outputExcel) {
        throw new Error("config must include inputExcel, sheetName, and outputExcel");
    }
    let maxRows = null;
    if (maxRowsRaw !== undefined && maxRowsRaw !== null && String(maxRowsRaw).trim() !== "") {
        const n = Number(maxRowsRaw);
        if (!Number.isFinite(n)) {
            throw new Error("config max_rows must be a number when provided");
        }
        if (n > 0) maxRows = Math.floor(n);
    }
    return { inputExcel, sheetName, outputExcel, maxRows, storageState };
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
const FILLED_OUTPUT_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_filled.csv`);

// ── Utility ──────────────────────────────────────────────────────────────────

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

function logStep(message, level = 0) {
    const indent = "  ".repeat(level);
    const ts = new Date().toISOString();
    console.log(`${indent}${ts} ${message}`);
}

function formatQty(v) {
    if (typeof v === "number") return v.toFixed(2);
    const s = cellText(v);
    const n = Number(s);
    if (Number.isFinite(n)) return n.toFixed(2);
    return s;
}

// ── Excel helpers ────────────────────────────────────────────────────────────

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

function buildEprSet(ws, headerMap) {
    const set = new Set();
    const col = headerMap.get(normHeader("EPR Invoice Number"));
    if (!col) return set;
    for (let r = 2; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        if (!row.hasValues) continue;
        const v = cellText(row.getCell(col).value);
        if (v) set.add(v);
    }
    return set;
}

// ── Logging ──────────────────────────────────────────────────────────────────

function ensureLogHeader() {
    if (fs.existsSync(LOG_PATH)) return;
    const header = [
        "datetime",
        "row",
        "gst_e_invoice_number",
        "registration_type",
        "entity_name",
        "quantity_tons",
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
        cellText(getVal(row, headerMap, "GST E-Invoice Number*")),
        cellText(getVal(row, headerMap, "Registration Type*")),
        cellText(getVal(row, headerMap, "Name of the Entity*")),
        cellText(getVal(row, headerMap, "Total Plastic Quantity (Tons)*")),
        cellText(eprInvoiceNumber),
        status,
        message || "",
    ].map(csvEscape);
    fs.appendFileSync(LOG_PATH, `${data.join(",")}\n`);
}

function ensureFilledHeader(headerList) {
    if (fs.existsSync(FILLED_OUTPUT_PATH)) return;
    const header = [...headerList, "datetime", "message"].map(csvEscape).join(",");
    fs.writeFileSync(FILLED_OUTPUT_PATH, `${header}\n`);
}

function appendFilledRow(row, headerMap, headerList, { message }) {
    ensureFilledHeader(headerList);
    const ts = new Date().toISOString();
    const rowData = headerList.map((h) => cellText(getVal(row, headerMap, h)));
    const data = [...rowData, ts, message || ""].map(csvEscape);
    fs.appendFileSync(FILLED_OUTPUT_PATH, `${data.join(",")}\n`);
}

// ── Safe Excel write ─────────────────────────────────────────────────────────

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

// ── Browser helpers ──────────────────────────────────────────────────────────

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

async function isLoginPage(page) {
    const pwd = page.locator('input[type="password"]').first();
    if (await pwd.count()) return true;
    const loginBtn = page.locator('button:has-text("Login"), button:has-text("Sign In")').first();
    if (await loginBtn.count()) return true;
    return false;
}

async function attemptLogout(page) {
    try {
        const logoutDirect = page.locator('text=/logout/i').first();
        if (await logoutDirect.count()) {
            await logoutDirect.click().catch(() => { });
            await page.waitForTimeout(1000);
            return true;
        }

        const toggles = [
            'button[aria-haspopup="menu"]',
            'button[aria-expanded]',
            '.dropdown-toggle',
            '.nav-link.dropdown-toggle',
            '.user-profile',
            '.profile',
        ];
        for (const sel of toggles) {
            const btn = page.locator(sel).first();
            if (await btn.count()) {
                await btn.click().catch(() => { });
                const logout = page.locator('text=/logout/i').first();
                if (await logout.count()) {
                    await logout.click().catch(() => { });
                    await page.waitForTimeout(1000);
                    return true;
                }
            }
        }
    } catch { }
    return false;
}

// ── Form interaction ─────────────────────────────────────────────────────────

async function fillByName(page, name, value) {
    const v = cellText(value);
    if (!v) return;
    const loc = page.locator(`input[name="${name}"]`).first();
    await loc.waitFor({ state: "visible", timeout: 30000 });
    await loc.scrollIntoViewIfNeeded();
    await loc.click();
    await loc.fill("");
    await loc.fill(v);
    await loc.blur();
}

async function selectNgSelectByLabel(page, labelText, optionText) {
    const text = cellText(optionText);
    if (!text) throw new Error(`Missing option for ${labelText}`);

    const group = page
        .locator(".form-group", { has: page.locator("label", { hasText: labelText }) })
        .first();

    logStep(`select ng-select: ${labelText} -> ${text}`, 1);
    await group.waitFor({ state: "visible", timeout: 20000 });
    const ng = group.locator("ng-select").first();
    await ng.scrollIntoViewIfNeeded();
    await ng.click();

    const panel = page.locator(".ng-dropdown-panel");
    await panel.waitFor({ state: "visible", timeout: 20000 });

    const searchInput = panel.locator("input[type='text']").first();
    if (await searchInput.count()) {
        try {
            await searchInput.fill(text);
            await page.waitForTimeout(120);
        } catch { }
    }

    const opt = panel.locator(".ng-option", { hasText: text }).first();
    await opt.waitFor({ state: "visible", timeout: 20000 });
    await opt.click();

    await panel.waitFor({ state: "hidden", timeout: 20000 }).catch(() => { });
}

async function pickEntityName(page, entityNameValue) {
    const name = cellText(entityNameValue);
    if (!name) throw new Error("Entity name empty");

    const group = page
        .locator(".form-group", { has: page.locator("label", { hasText: "Name of the Entity" }) })
        .first();

    if (await group.count()) {
        const ng = group.locator("ng-select").first();
        if (await ng.count()) {
            logStep("pick entity name: ng-select", 1);
            await ng.scrollIntoViewIfNeeded();
            await ng.click();

            const panel = page.locator(".ng-dropdown-panel");
            await panel.waitFor({ state: "visible", timeout: 20000 });

            const searchInput = panel.locator("input[type='text']").first();
            if (await searchInput.count()) {
                await searchInput.fill(name);
                await page.waitForTimeout(150);
                await searchInput.press("Enter");
            } else {
                await page.keyboard.type(name);
                await page.waitForTimeout(300);
                await page.keyboard.press("Enter");
            }

            await panel.waitFor({ state: "hidden", timeout: 5000 }).catch(() => { });
            const selected = group.locator(".ng-value-label").first();
            if (await selected.count()) {
                const selectedText = (await selected.innerText().catch(() => "")).trim();
                if (selectedText) {
                    logStep(`pick entity name: selected "${selectedText}"`, 1);
                    await waitForLoaderToFinish(page);
                    return;
                }
            }

            logStep("pick entity name: enter did not select, fallback to click", 1);
            await selectNgSelectByLabel(page, "Name of the Entity", name);
            await waitForLoaderToFinish(page);
            return;
        }

        // Fallback: plain input (for unregistered entities)
        const input = group.locator("input").first();
        if (await input.count()) {
            logStep("pick entity name: text input", 1);
            await input.waitFor({ state: "visible", timeout: 30000 });
            await input.scrollIntoViewIfNeeded();
            await input.click();
            await input.fill(name);
            await input.blur();
            await waitForLoaderToFinish(page);
            return;
        }
    }

    throw new Error("Name of the Entity field not found on form");
}

async function waitEntityAutofill(page) {
    await waitForLoaderToFinish(page);
    const addr = page.locator('input[name="address"]').first();
    if (!(await addr.count())) return;
    await addr.waitFor({ state: "visible", timeout: 10000 }).catch(() => { });
    // Give auto-fill a chance to populate
    await page.waitForTimeout(500);
}

// ── Submit & EPR read ────────────────────────────────────────────────────────

async function clickSubmitAndConfirm(page) {
    const submit = page.locator('button[type="submit"]', { hasText: "Generate EPR Invoice Number" }).first();
    logStep("submit: start", 1);
    await submit.waitFor({ state: "visible", timeout: 20000 });
    await submit.scrollIntoViewIfNeeded();
    if (await submit.isDisabled()) {
        throw new Error("Submit disabled: some required fields still missing.");
    }
    await submit.click();
    try {
        const confirmBtn = page
            .locator("#openConfirmation .modal-footer button.btn-primary", { hasText: "Confirm" })
            .first();
        await confirmBtn.waitFor({ state: "visible", timeout: 60000 });
        await confirmBtn.click();
    } catch { }
    logStep("submit: done", 1);
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
        if (/confirm entered details/i.test(text)) return "";
        const match = text.match(/EPR\s*Invoice\s*Number\s*[:\-]?\s*([A-Za-z0-9\-\/]+)/i);
        if (match && match[1]) return match[1].trim();
        const sibling = label.locator("xpath=following::span[1] | following::div[1] | following::p[1]").first();
        const sibText = (await sibling.innerText().catch(() => "")).trim();
        if (/confirm entered details/i.test(sibText)) return "";
        if (sibText) return sibText;
    }
    return "";
}

async function waitForEprInvoiceNumber(page, timeoutMs = 20000) {
    const start = Date.now();
    logStep("wait EPR invoice: start", 1);
    while (Date.now() - start < timeoutMs) {
        const modal = page.locator(".modal-dialog, .modal-content").first();
        if (await modal.count()) {
            const confirmBtn = modal.getByRole("button", { name: "Confirm", exact: true }).first();
            if (await confirmBtn.count()) {
                await confirmBtn.click().catch(() => { });
                await modal.waitFor({ state: "hidden", timeout: 5000 }).catch(() => { });
            }
        }
        const val = await readEprInvoiceNumber(page);
        if (val) return val;
        await page.waitForTimeout(300);
    }
    logStep("wait EPR invoice: timeout", 1);
    return "";
}

// ── Navigation / reset ───────────────────────────────────────────────────────

async function clickAddNew(page) {
    const addNewBtn = page.getByRole("button", { name: "Add New", exact: true }).first();
    logStep("click add new: start", 1);
    await addNewBtn.waitFor({ state: "visible", timeout: 60000 });
    await addNewBtn.scrollIntoViewIfNeeded();
    await addNewBtn.click();
    await page.waitForTimeout(300);
    logStep("click add new: done", 1);
}

async function clickAddNewIfVisible(page) {
    const addNewBtn = page.getByRole("button", { name: "Add New", exact: true }).first();
    if (!(await addNewBtn.count())) return false;
    await addNewBtn.waitFor({ state: "visible", timeout: 5000 }).catch(() => { });
    await addNewBtn.scrollIntoViewIfNeeded().catch(() => { });
    await addNewBtn.click().catch(() => { });
    await page.waitForTimeout(300);
    logStep("click add new (visible): done", 2);
    return true;
}

async function closeModalIfOpen(page) {
    try {
        const modal = page.locator(".modal-dialog").first();
        if (!(await modal.isVisible().catch(() => false))) return;

        const closeBtn = modal.locator("button.close, button[data-bs-dismiss='modal']").first();
        if (await closeBtn.count()) {
            await closeBtn.click().catch(() => { });
            await modal.waitFor({ state: "hidden", timeout: 5000 }).catch(() => { });
            logStep("modal closed", 2);
        }
    } catch { }
}

async function resetToFreshPage(page) {
    logStep("reset to fresh page: start", 1);
    await closeModalIfOpen(page);
    await page.goto(PIBO_URL, { waitUntil: "domcontentloaded" }).catch(() => { });
    await page.waitForTimeout(2000);
    await waitForLoaderToFinish(page);
    logStep("reset to fresh page: done", 1);
}

// ── Main ─────────────────────────────────────────────────────────────────────

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
    const eprSet = buildEprSet(ws, headerMap);
    const headerList = getHeaderList(ws);

    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext(
        fs.existsSync(STORAGE) ? { storageState: STORAGE } : {}
    );
    const page = await context.newPage();

    await page.goto(PIBO_URL, { waitUntil: "domcontentloaded" });
    if (!fs.existsSync(STORAGE) || (await isLoginPage(page))) {
        await attemptLogout(page);
        console.log("Login manually in this Playwright window, then press ENTER here...");
        await new Promise((res) => process.stdin.once("data", () => res()));
        await context.storageState({ path: STORAGE });
        console.log("Saved session to storageState_pibo.json");
    }

    await page.goto(PIBO_URL, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(3000);
    await waitForLoaderToFinish(page);

    const lastRow = CONFIG.maxRows ? Math.min(ws.rowCount, CONFIG.maxRows) : ws.rowCount;
    for (let r = 2; r <= lastRow; r++) {
        const row = ws.getRow(r);
        let successThisRow = false;
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

        // Read all fields from Excel
        const regType = getVal(row, headerMap, "Registration Type*");
        const entityType = getVal(row, headerMap, "Entity Type*");
        const entityName = getVal(row, headerMap, "Name of the Entity*");
        const address = getVal(row, headerMap, "Address*");
        const state = getVal(row, headerMap, "State*");
        const mobile = getVal(row, headerMap, "Mobile Number*");
        const plasticMaterial = getVal(row, headerMap, "Plastic Material Type*");
        const plasticCategory = getVal(row, headerMap, "Category of Plastic*");
        const financialYear = getVal(row, headerMap, "Financial Year*");
        const gst = getVal(row, headerMap, "GST*");
        const bankAccount = getVal(row, headerMap, "Bank Account No*");
        const ifsc = getVal(row, headerMap, "IFSC Code*");
        const gstPaid = getVal(row, headerMap, "GST Paid*");
        const gstInvoice = getVal(row, headerMap, "GST E-Invoice Number*");
        const quantity = getVal(row, headerMap, "Total Plastic Quantity (Tons)*");
        const recycledPlastic = getVal(row, headerMap, "% of Recycled Plastic Content*");

        try {
            console.log(`Row ${r} starting...`);

            // Open form modal
            await clickAddNew(page);
            await page.waitForTimeout(500);
            await waitForLoaderToFinish(page);

            // 1. Registration Type
            logStep("fill Registration Type", 1);
            await selectNgSelectByLabel(page, "Registration Type", regType);
            await page.waitForTimeout(300);

            // 2. Entity Type
            logStep("fill Entity Type", 1);
            await selectNgSelectByLabel(page, "Entity Type", entityType);
            await page.waitForTimeout(500);
            await waitForLoaderToFinish(page);

            // 3. Name of the Entity (appears after Entity Type is selected)
            logStep("fill Name of the Entity", 1);
            await pickEntityName(page, entityName);
            await waitEntityAutofill(page);

            // 4. Address
            logStep("fill Address", 1);
            await fillByName(page, "address", address);

            // 5. State
            logStep("fill State", 1);
            await selectNgSelectByLabel(page, "State", state);

            // 6. Mobile Number
            logStep("fill Mobile Number", 1);
            await fillByName(page, "mobile_number", mobile);

            // 7. Plastic Material Type
            logStep("fill Plastic Material Type", 1);
            await selectNgSelectByLabel(page, "Plastic Material Type", plasticMaterial);
            await page.waitForTimeout(300);

            // 8. Category of Plastic
            logStep("fill Category of Plastic", 1);
            await selectNgSelectByLabel(page, "Category of Plastic", plasticCategory);

            // 9. Financial Year
            logStep("fill Financial Year", 1);
            await selectNgSelectByLabel(page, "Financial Year", financialYear);

            // 10. GST
            logStep("fill GST", 1);
            await fillByName(page, "gst_no", gst);

            // 11. Bank Account No
            logStep("fill Bank Account No", 1);
            await fillByName(page, "account_no", bankAccount);

            // 12. IFSC Code
            logStep("fill IFSC Code", 1);
            await fillByName(page, "ifsc", ifsc);

            // 13. GST Paid / Total GST Paid
            logStep("fill GST Paid", 1);
            await fillByName(page, "gst_paid", gstPaid);

            // 14. GST E-Invoice Number
            logStep("fill GST E-Invoice Number", 1);
            await fillByName(page, "gst_invoice", gstInvoice);

            // 15. Total Plastic Quantity (Tons)
            logStep("fill Quantity", 1);
            await fillByName(page, "quantity", formatQty(quantity));

            // 16. % of Recycled Plastic Content
            logStep("fill Recycled Plastic %", 1);
            await fillByName(page, "recycled_plastic", recycledPlastic);

            // Submit and confirm
            await clickSubmitAndConfirm(page);
            logStep("post-submit: wait", 1);
            await page.waitForTimeout(300);
            await waitForLoaderToFinish(page);
            const toastText = await readToastTextWithRetry(page);
            if (toastText) {
                logStep(`toast: ${toastText}`, 1);
            }

            const eprInvoice = await waitForEprInvoiceNumber(page);
            if (eprSet.has(eprInvoice)) {
                throw new Error("Duplicate EPR Invoice Number: " + eprInvoice);
            }
            if (!eprInvoice) {
                throw new Error("EPR Invoice Number not found after submit.");
            }
            logStep(`EPR invoice: ${eprInvoice}`, 1);

            setVal(row, headerMap, "Status", "Filled");
            setVal(row, headerMap, "EPR Invoice Number", eprInvoice);
            eprSet.add(eprInvoice);
            row.commit();
            await safeWriteWorkbook(wb);
            await syncInputWorkbook(wb);

            appendLogRow(row, headerMap, {
                status: "Filled",
                eprInvoiceNumber: eprInvoice,
                message: toastText,
            });
            appendFilledRow(row, headerMap, headerList, {
                message: toastText,
            });

            console.log(`Row ${r}: Filled`);
            successThisRow = true;
        } catch (e) {
            const msg = String(e?.message || e);
            console.log(`Row ${r}: Failed ->`, msg);
            setVal(row, headerMap, "Status", "Failed: " + msg);
            row.commit();
            await safeWriteWorkbook(wb);
            await syncInputWorkbook(wb);
            appendLogRow(row, headerMap, {
                status: "Failed",
                eprInvoiceNumber: "",
                message: msg,
            });
            appendFilledRow(row, headerMap, headerList, {
                message: msg,
            });
        } finally {
            if (page.isClosed()) {
                console.log("Page closed. Stopping.");
                break;
            }
            if (successThisRow) {
                await page.waitForTimeout(300);
            }
            await waitForLoaderToFinish(page);
            await closeModalIfOpen(page);
            await page.waitForTimeout(500);
        }
    }

    await browser.close();
    console.log("Done. Updated Excel:", EXCEL_PATH);
})();
