const ExcelJS = require("exceljs");
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

const URL = "https://eprplastic.cpcb.gov.in/#/epr/details/sales";
const CONFIG_PATH = path.resolve(__dirname, "config.json");
const STORAGE = path.resolve(__dirname, "storageState.json");

function loadConfig() {
    if (!fs.existsSync(CONFIG_PATH)) {
        throw new Error(`Missing config file: ${CONFIG_PATH}`);
    }
    let raw = "";
    try {
        raw = fs.readFileSync(CONFIG_PATH, "utf8");
    } catch (e) {
        throw new Error(`Failed to read config file: ${e?.message || e}`);
    }
    let cfg = null;
    try {
        cfg = JSON.parse(raw);
    } catch (e) {
        throw new Error(`Invalid JSON in config file: ${e?.message || e}`);
    }
    const inputExcel = String(cfg?.inputExcel || "").trim();
    const sheetName = String(cfg?.sheetName || "").trim();
    const outputExcel = String(cfg?.outputExcel || "").trim();
    if (!inputExcel || !sheetName || !outputExcel) {
        throw new Error("config.json must include inputExcel, sheetName, and outputExcel");
    }
    return { inputExcel, sheetName, outputExcel };
}

const CONFIG = loadConfig();
const EXCEL_PATH = path.resolve(__dirname, CONFIG.inputExcel);
const SHEET = CONFIG.sheetName;
const OUTPUT_PATH = path.resolve(__dirname, CONFIG.outputExcel);
const EXCEL_TMP = `${OUTPUT_PATH}.tmp`;
const EXCEL_BAK = `${OUTPUT_PATH}.bak`;
const OUTPUT_BASENAME = path.basename(OUTPUT_PATH, path.extname(OUTPUT_PATH));
const LOG_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_log.csv`);
const FILLED_OUTPUT_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_filled.csv`);

// ---------- Helpers ----------
function normHeader(s) {
    return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object" && v.text) return String(v.text).trim();
    return String(v).trim();
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

function safeSalesDateString(v) {
    try {
        return excelDateToISO(v);
    } catch {
        return cellText(v);
    }
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
        "registration_type",
        "entity_name",
        "seller_gst",
        "buyer_gst",
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
        safeSalesDateString(getVal(row, headerMap, "Sales date*")),
        cellText(getVal(row, headerMap, "Quantity Sold(MT)")),
        cellText(getVal(row, headerMap, "Registration Type*")),
        cellText(getVal(row, headerMap, "Name of the Entity *")),
        cellText(getVal(row, headerMap, "GST No. of Seller *")),
        cellText(getVal(row, headerMap, "Buyer GST")),
        cellText(eprInvoiceNumber),
        status,
        message || "",
    ].map(csvEscape);
    fs.appendFileSync(LOG_PATH, `${data.join(",")}\n`);
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

async function fillById(page, id, value) {
    const v = cellText(value);
    const loc = page.locator(`#${id}`);
    await loc.waitFor({ state: "visible", timeout: 60000 });
    await loc.scrollIntoViewIfNeeded();
    await loc.click({ timeout: 15000 });
    await loc.fill("");
    if (v) await loc.fill(v);
    await loc.blur();
}

async function fillBySelector(page, selector, value) {
    const v = cellText(value);
    const loc = page.locator(selector).first();
    await loc.waitFor({ state: "visible", timeout: 60000 });
    await loc.scrollIntoViewIfNeeded();
    await loc.click({ timeout: 15000 });
    await loc.fill("");
    if (v) await loc.fill(v);
    await loc.blur();
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

function getValAny(row, headerMap, headerNames) {
    for (const name of headerNames) {
        const v = getVal(row, headerMap, name);
        if (cellText(v)) return v;
    }
    return "";
}

function setVal(row, headerMap, headerName, value) {
    const col = headerMap.get(normHeader(headerName));
    if (!col) return;
    row.getCell(col).value = value;
}

function setValAny(row, headerMap, headerNames, value) {
    for (const name of headerNames) {
        const col = headerMap.get(normHeader(name));
        if (col) {
            row.getCell(col).value = value;
            return true;
        }
    }
    return false;
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
    // Write to temp, then replace output to avoid 0-byte corruption on crash.
    await wb.xlsx.writeFile(EXCEL_TMP);
    try {
        if (fs.existsSync(OUTPUT_PATH) && fs.statSync(OUTPUT_PATH).size > 0) {
            fs.copyFileSync(OUTPUT_PATH, EXCEL_BAK);
        }
    } catch { }
    fs.renameSync(EXCEL_TMP, OUTPUT_PATH);
}

// ---------- ERP helpers ----------

// Try to wait for common loaders to disappear.
// If no loader exists, this just continues.
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

// Step 0: Click Add New
async function clickAddNew(page) {
    const addNewBtn = page.getByRole("button", { name: "Add New", exact: true }).first();
    await addNewBtn.waitFor({ state: "visible", timeout: 60000 });
    await addNewBtn.scrollIntoViewIfNeeded();
    await addNewBtn.click();
    await page.waitForTimeout(300); // small UI settle
}

// Step 1: Select CAT-II row checkbox in top table (optionally by Plastic Type)
async function selectCat2Row(page, plasticTypeText) {
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

    await catRow.waitFor({ state: "visible", timeout: 20000 });

    const checkbox = catRow.locator('input[type="checkbox"][name="check-box"]').first();
    await checkbox.scrollIntoViewIfNeeded();
    await checkbox.click({ force: true });

    // After selecting CAT-II, the form fields appear
    await page.waitForSelector('input[name="qty_product_sold"]', { timeout: 30000 });
}

// Step 2: ng-select by label text (no need formcontrolname)
async function selectNgSelectByLabel(page, labelText, optionText) {
    const text = cellText(optionText);
    if (!text) throw new Error(`Missing option for ${labelText}`);

    const group = page
        .locator(".form-group", { has: page.locator("label", { hasText: labelText }) })
        .first();

    await group.waitFor({ state: "visible", timeout: 20000 });

    const ng = group.locator("ng-select").first();
    await ng.scrollIntoViewIfNeeded();
    await ng.click();

    const panel = page.locator(".ng-dropdown-panel");
    await panel.waitFor({ state: "visible", timeout: 20000 });

    // If there's search input, type to filter
    const searchInput = panel.locator("input[type='text']").first();
    if (await searchInput.count()) {
        try {
            await searchInput.fill(text);
            await page.waitForTimeout(200);
        } catch { }
    }

    const opt = panel.locator(".ng-option", { hasText: text }).first();
    await opt.waitFor({ state: "visible", timeout: 20000 });
    await opt.click();

    await panel.waitFor({ state: "hidden", timeout: 20000 }).catch(() => { });
}

async function clickSubmitAndConfirm(page) {
    const submit = page.locator('button[type="submit"]', { hasText: "Generate EPR Invoice Number" }).first();
    await submit.waitFor({ state: "visible", timeout: 20000 });
    if (await submit.isDisabled()) {
        throw new Error("Submit disabled: some required fields still missing.");
    }
    await submit.click();

    // Confirm modal (expected after submit)
    try {
        const confirmBtn = page.locator(".modal-footer button", { hasText: "Confirm" }).first();
        await confirmBtn.waitFor({ state: "visible", timeout: 60000 });
        await confirmBtn.click();
    } catch { }
}

async function clickResetAndConfirm(page) {
    const reset = page.locator("button", { hasText: /\\bReset\\b/i }).first();
    await reset.waitFor({ state: "visible", timeout: 20000 });
    await reset.scrollIntoViewIfNeeded();
    await reset.click();

    const modal = page.locator(".modal-dialog, .modal-content").first();
    if (await modal.count()) {
        try {
            await modal.waitFor({ state: "visible", timeout: 15000 });
            const confirmBtn = modal.getByRole("button", { name: "Confirm", exact: true }).first();
            if (await confirmBtn.count()) {
                await confirmBtn.click();
            }
        } catch { }
    }
}

async function waitForToast(page) {
    const toast = page.locator(".toast, .toaster, .ngx-toastr, .toast-container").first();
    if (await toast.count()) {
        await toast.waitFor({ state: "visible", timeout: 20000 }).catch(() => { });
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

async function readEprInvoiceNumber(page) {
    // Preferred: read disabled input value
    const input = page.locator("#invoiceNumberCopy").first();
    if (await input.count()) {
        try {
            const val = (await input.inputValue()).trim();
            if (val) return val;
        } catch { }
    }

    // Try common patterns: label + value near it.
    const label = page.locator("text=/EPR\\s*Invoice\\s*Number/i").first();
    if (await label.count()) {
        const container = label.locator("xpath=ancestor-or-self::*[self::div or self::span or self::p][1]");
        const text = (await container.innerText().catch(() => "")) || "";
        const match = text.match(/EPR\s*Invoice\s*Number\s*[:\-]?\s*([A-Za-z0-9\-\/]+)/i);
        if (match && match[1]) return match[1].trim();
        const sibling = label.locator("xpath=following::span[1] | following::div[1] | following::p[1]").first();
        const sibText = (await sibling.innerText().catch(() => "")).trim();
        if (sibText) return sibText;
    }

    // Fallback: look for copy button near bottom-left block
    const copyBlock = page.locator("button", { hasText: /copy/i }).first();
    if (await copyBlock.count()) {
        const parent = copyBlock.locator("xpath=ancestor::*[self::div or self::p][1]");
        const text = (await parent.innerText().catch(() => "")).trim();
        const match = text.match(/([A-Za-z0-9\-\/]{6,})/);
        if (match && match[1]) return match[1].trim();
    }
    return "";
}

// Step 3: Entity Name autocomplete (handles both patterns)
async function pickEntityName(page, entityNameValue) {
    const name = cellText(entityNameValue);
    if (!name) throw new Error("Entity name empty");

    const input = page.locator('input[formcontrolname="entity_name"]').first();
    await input.waitFor({ state: "visible", timeout: 30000 });

    await input.scrollIntoViewIfNeeded();
    await input.click();
    await input.fill(name);

    // Wait a little for autocomplete
    await page.waitForTimeout(800);

    // Generic dropdown patterns (not ng-select)
    const suggestion = page.locator(
        'ul li, .dropdown-item, .typeahead-item, .autocomplete-items div'
    ).first();

    if (await suggestion.count()) {
        await suggestion.click();
    } else {
        // If no suggestion appears, just blur
        await input.blur();
    }

    // Wait for autofill loader
    await waitForLoaderToFinish(page);
}


// Step 4: wait address autofill (after selecting entity)
async function waitEntityAutofill(page) {
    await waitForLoaderToFinish(page);

    const addr = page.locator('input[formcontrolname="entity_address"]').first();
    await addr.waitFor({ state: "visible", timeout: 20000 });

    // Wait until address becomes non-empty (autofill)
    await page
        .waitForFunction(() => {
            const a = document.querySelector('input[formcontrolname="entity_address"]');
            return a && a.value && a.value.trim().length > 3;
        }, { timeout: 30000 })
        .catch(() => { });
}

// ---------- Main ----------
(async () => {
    // Safety check: file exists + not empty
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
        console.log("Sheets found:");
        wb.worksheets.forEach((x) => console.log(" -", x.name));
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

    // login once (only if no storage state)
    if (!fs.existsSync(STORAGE)) {
        console.log("Login manually in this Playwright window, then press ENTER here...");
        await new Promise((res) => process.stdin.once("data", () => res()));
        await context.storageState({ path: STORAGE });
        console.log("Saved session to storageState.json");
    }

    // ensure we are on sales page and table exists
    await page.goto(URL, { waitUntil: "domcontentloaded" });
    await page.waitForSelector("#ScrollableSimpleTableBody", { timeout: 60000 });
    await clickAddNew(page);
    for (let r = 2; r <= ws.rowCount; r++) {
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

        // Read values
        const qtySold = getVal(row, headerMap, "Quantity Sold(MT)");
        const regType = getVal(row, headerMap, "Registration Type*");
        const entityType = getVal(row, headerMap, "Entity Type*");
        const entityName = getVal(row, headerMap, "Name of the Entity *");
        const entityAddress = getVal(row, headerMap, "Address*");
        const entityState = getVal(row, headerMap, "State*");
        const entityDistrict = getVal(row, headerMap, "District*");

        const sellerGst = getVal(row, headerMap, "GST No. of Seller *");
        const buyerGst = getVal(row, headerMap, "Buyer GST");
        const hsn = getVal(row, headerMap, "HSN CODE");
        const invno = getVal(row, headerMap, "E-Invoice Number*");
        const account = getVal(row, headerMap, "Bank Account No*");
        const ifsc = getVal(row, headerMap, "IFSC Code*");
        const principal = getVal(row, headerMap, "Principal Amount(₹)*");
        const gstOther = getVal(row, headerMap, "GST & Other Charges(₹)*");
        const salesDateRaw = getVal(row, headerMap, "Sales date*");

        try {
            console.log(`Row ${r} starting...`);

            // âœ… Then select CAT-II checkbox (to reveal forms)
            await selectCat2Row(page, "PP");

            // âœ… Qty Sold
            await fillBySelector(page, 'input[name="qty_product_sold"]', qtySold.toFixed(4));

            // âœ… Registration Type
            await selectNgSelectByLabel(page, "Registration Type", regType);

            const regTypeNorm = cellText(regType).toLowerCase();
            if (regTypeNorm === "registered") {
                // âœ… Entity Type (only for Registered)
                await selectNgSelectByLabel(page, "Entity Type", entityType);
                await waitForLoaderToFinish(page);

                // âœ… Entity Name search + select (autocomplete)
                await pickEntityName(page, entityName);

                // âœ… Wait for autofill
                await waitEntityAutofill(page);
            } else if (regTypeNorm === "unregistered") {
                // âœ… Unregistered: manual entity fields, no entity-type wait
                await fillBySelector(page, 'input[formcontrolname="entity_name"]', entityName);
                await fillBySelector(page, 'input[formcontrolname="entity_address"]', entityAddress);
                await selectNgSelectByLabel(page, "State", entityState);
                await selectNgSelectByLabel(page, "District", entityDistrict);
            } else {
                throw new Error(`Unsupported Registration Type: ${cellText(regType)}`);
            }

            // âœ… Fill remaining fields
            await fillById(page, "sellerGst", sellerGst);
            await fillById(page, "buyerGst", buyerGst);
            await fillById(page, "hsnCode", hsn);
            await fillById(page, "invno", invno);
            await fillById(page, "account_number", account);
            await fillById(page, "ifsc_code", ifsc);
            await fillById(page, "amount", principal);
            await fillById(page, "gst", gstOther);

            // âœ… Sales date (keyboard blocked -> JS set)
            const salesDateISO = excelDateToISO(salesDateRaw);
            await setAngularDateById(page, "salesDate", salesDateISO);


            // âœ… Submit and confirm
            await clickSubmitAndConfirm(page);
            await page.waitForTimeout(1000);
            await waitForToast(page);
            const toastText = await readToastText(page);
            if (toastText) {
                console.log(`Row ${r}: Toast -> ${toastText}`);
            }


            const eprInvoice = await readEprInvoiceNumber(page);
            console.log(eprInvoice)
            if (!eprInvoice) {
                throw new Error("EPR Invoice Number not found after submit.");
            }

            // âœ… Update Excel status
            setVal(row, headerMap, "Status", "Filled");
            setVal(row, headerMap, "EPR Invoice Number", eprInvoice);
            row.commit();
            await safeWriteWorkbook(wb);

            appendLogRow(row, headerMap, {
                status: "Filled",
                eprInvoiceNumber: eprInvoice,
                message: toastText,
            });
            appendFilledRow(row, headerMap, headerList, {
                message: toastText,
            });

            console.log(`Row ${r}: Filled âœ…`);
            await page.waitForTimeout(8000);

        } catch (e) {
            const msg = String(e?.message || e);
            console.log(`Row ${r}: Failed âŒ ->`, msg);

            setVal(row, headerMap, "Status", "Failed: " + msg);
            row.commit();
            await safeWriteWorkbook(wb);

            appendLogRow(row, headerMap, {
                status: "Failed",
                eprInvoiceNumber: "",
                message: msg,
            });
        }
    }

    await browser.close();
    console.log("Done. Updated Excel:", EXCEL_PATH);
})();





