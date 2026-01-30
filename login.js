const { chromium } = require("playwright");
const path = require("path");

const STORAGE = path.resolve(__dirname, "storageState.json");
const URL = "https://eprplastic.cpcb.gov.in/#/";

(async () => {
    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext();
    const page = await context.newPage();

    await page.goto(URL, { waitUntil: "domcontentloaded" });

    console.log("✅ Login manually in this Playwright window.");
    console.log("✅ After login, go to Sales page once.");
    console.log("✅ Then come back here and press ENTER.");

    await new Promise((res) => process.stdin.once("data", res));

    await context.storageState({ path: STORAGE });
    console.log("✅ Saved:", STORAGE);

    await browser.close();
})();
