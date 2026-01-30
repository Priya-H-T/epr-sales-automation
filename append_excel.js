const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const cfg = require("./config.json");
const inputPath = path.resolve(__dirname, cfg.inputExcel);
const outputPath = path.resolve(__dirname, cfg.outputExcel);
const sheetName = cfg.sheetName;

async function appendOutputToInput() {
    if (!fs.existsSync(inputPath)) {
        throw new Error(`Input not found: ${inputPath}`);
    }
    if (!fs.existsSync(outputPath)) {
        throw new Error(`Output not found: ${outputPath}`);
    }

const wbIn = new ExcelJS.Workbook();
const wbOut = new ExcelJS.Workbook();
    await wbIn.xlsx.readFile(inputPath);
    await wbOut.xlsx.readFile(outputPath);

    const wsIn = wbIn.getWorksheet(sheetName);
    const wsOut = wbOut.getWorksheet(sheetName);
    if (!wsIn || !wsOut) {
        throw new Error(`Sheet not found: ${sheetName}`);
    }

    const normHeader = (s) => String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
    const cellText = (v) => {
        if (v === null || v === undefined) return "";
        if (typeof v === "object" && v.text) return String(v.text).trim();
        return String(v).trim();
    };

    const getHeaderMap = (ws) => {
        const map = new Map();
        const headerRow = ws.getRow(1);
        headerRow.eachCell((cell, colNumber) => {
            const key = normHeader(cellText(cell.value));
            if (key) map.set(key, colNumber);
        });
        return map;
    };

    const inHeaderMap = getHeaderMap(wsIn);
    const outHeaderMap = getHeaderMap(wsOut);

    const eprKey = normHeader("EPR Invoice Number");
    const einvKey = normHeader("E-Invoice Number*");

    const idxInEpr = inHeaderMap.get(eprKey);
    const idxInEinv = inHeaderMap.get(einvKey);
    const idxOutEpr = outHeaderMap.get(eprKey);
    const idxOutEinv = outHeaderMap.get(einvKey);

    if (!idxInEpr && !idxInEinv) {
        throw new Error("Input missing both EPR Invoice Number and E-Invoice Number* columns.");
    }
    if (!idxOutEpr && !idxOutEinv) {
        throw new Error("Output missing both EPR Invoice Number and E-Invoice Number* columns.");
    }

    const buildIndex = (ws, idx) => {
        const map = new Map();
        if (!idx) return map;
        for (let r = 2; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            if (!row.hasValues) continue;
            const key = cellText(row.getCell(idx).value);
            if (!key) continue;
            if (!map.has(key)) map.set(key, r);
        }
        return map;
    };

    const inByEpr = buildIndex(wsIn, idxInEpr);
    const inByEinv = buildIndex(wsIn, idxInEinv);

    // Update rows in input using output values
    for (let r = 2; r <= wsOut.rowCount; r++) {
        const rowOut = wsOut.getRow(r);
        if (!rowOut.hasValues) continue;
        const eprVal = idxOutEpr ? cellText(rowOut.getCell(idxOutEpr).value) : "";
        const einvVal = idxOutEinv ? cellText(rowOut.getCell(idxOutEinv).value) : "";

        let targetRowNum = null;
        if (eprVal && inByEpr.has(eprVal)) {
            targetRowNum = inByEpr.get(eprVal);
        } else if (einvVal && inByEinv.has(einvVal)) {
            targetRowNum = inByEinv.get(einvVal);
        }

        if (!targetRowNum) {
            continue;
        }

        const rowIn = wsIn.getRow(targetRowNum);
        for (const [header, outCol] of outHeaderMap.entries()) {
            const inCol = inHeaderMap.get(header);
            if (!inCol) continue;
            rowIn.getCell(inCol).value = rowOut.getCell(outCol).value;
        }
        rowIn.commit();
    }

    const tmp = `${inputPath}.tmp`;
    const bak = `${inputPath}.bak`;
    await wbIn.xlsx.writeFile(tmp);
    try {
        if (fs.existsSync(inputPath) && fs.statSync(inputPath).size > 0) {
            fs.copyFileSync(inputPath, bak);
        }
    } catch { }
    fs.renameSync(tmp, inputPath);
}

appendOutputToInput()
    .then(() => {
        console.log("Append complete. Input updated:", inputPath);
    })
    .catch((err) => {
        console.error("Append failed:", err.message || err);
        process.exit(1);
    });
