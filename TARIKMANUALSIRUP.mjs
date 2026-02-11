// TARIKMANUALSIRUP.mjs
// Gabungan: SIRUP PENYEDIA + SIRUP SWAKELOLA -> 2 sheet dalam 1 spreadsheet
//
// Spreadsheet: 1kp6tPQs18YfzBqUvnvsdQyps2VTy42Hc_qwbUgaITQ0
// Key: inaproc2026key.json
//
// Output sheets:
// - TARIKMANUALPENYEDIA
// - TARIKMANUALSWAKELOLA
//
// Install (sekali):
//   npm i playwright googleapis
//   npx playwright install chromium
//
// Run:
//   node TARIKMANUALSIRUP.mjs

import { chromium } from "playwright";
import { google } from "googleapis";
import fs from "fs";
import path from "path";

// =====================
// KONFIG
// =====================
const ID_KLDI = "D101";
const TAHUN = "2026";

const SPREADSHEET_ID = "1kp6tPQs18YfzBqUvnvsdQyps2VTy42Hc_qwbUgaITQ0";
const KEY_FILE = "inaproc2026key.json";

const SHEET_PENYEDIA = "TARIKMANUALPENYEDIA";
const SHEET_SWAKELOLA = "TARIKMANUALSWAKELOLA";

// performa
const PAGE_SIZE = 100;
const SLEEP_MS = 350;
const FLUSH_EVERY = 1000;

// endpoints
const ENDPOINT_PENYEDIA =
  "https://sirup.inaproc.id/sirup/datatablectr/dataruppenyediakldi";
const ENDPOINT_SWAKELOLA =
  "https://sirup.inaproc.id/sirup/datatablectr/datarupswakelolakldi";

// referer (cookie/session)
const REFERER_PENYEDIA =
  `https://sirup.inaproc.id/sirup/rekap/penyedia/${ID_KLDI}?tahun=${TAHUN}`;
const REFERER_SWAKELOLA =
  `https://sirup.inaproc.id/sirup/rekap/swakelola/${ID_KLDI}?tahun=${TAHUN}`;

// =====================
// HELPERS
// =====================
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function loadCredentials() {
  const keyPath = path.resolve(process.cwd(), KEY_FILE);
  if (!fs.existsSync(keyPath)) {
    throw new Error(
      `File key tidak ketemu: ${keyPath}\n` +
      `Pastikan ${KEY_FILE} ada di folder: ${process.cwd()}`
    );
  }
  const creds = JSON.parse(fs.readFileSync(keyPath, "utf8"));
  if (creds.type !== "service_account") {
    throw new Error(`Key salah. "type" harus "service_account".`);
  }
  if (!creds.client_email || !creds.private_key) {
    throw new Error(`Key kurang. Harus ada client_email & private_key.`);
  }
  creds.private_key = String(creds.private_key).replace(/\\n/g, "\n");
  return creds;
}

function toNumber(v) {
  if (v == null) return "";
  const s = String(v).replace(/[^\d.-]/g, "");
  if (!s) return "";
  const n = Number(s);
  return Number.isFinite(n) ? n : "";
}

// =====================
// GOOGLE SHEETS
// =====================
async function getSheetsClient() {
  const credentials = loadCredentials();
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const client = await auth.getClient();
  return google.sheets({ version: "v4", auth: client });
}

async function ensureSheet(sheets, title) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const exists = (meta.data.sheets || []).some(
    (s) => s.properties?.title === title
  );

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title } } }],
      },
    });
  }

  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${title}!A:Z`,
  });
}

async function writeHeader(sheets, sheetName, header) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1:H1`,
    valueInputOption: "RAW",
    requestBody: { values: [header] },
  });
}

async function appendRows(sheets, sheetName, rows) {
  if (!rows.length) return;
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A:H`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows },
  });
}

// =====================
// URL BUILDERS (FULL PARAMS)
// =====================
function buildUrlPenyedia({ sEcho, start, length }) {
  const params = {
    idKldi: ID_KLDI,
    tahun: TAHUN,
    sEcho: String(sEcho),
    iColumns: "8",
    sColumns: ",,namaPaket,,,,,",
    iDisplayStart: String(start),
    iDisplayLength: String(length),

    mDataProp_0: "0",
    sSearch_0: "",
    bRegex_0: "false",
    bSearchable_0: "true",
    bSortable_0: "false",

    mDataProp_1: "1",
    sSearch_1: "",
    bRegex_1: "false",
    bSearchable_1: "true",
    bSortable_1: "true",

    mDataProp_2: "2",
    sSearch_2: "",
    bRegex_2: "false",
    bSearchable_2: "true",
    bSortable_2: "true",

    mDataProp_3: "3",
    sSearch_3: "",
    bRegex_3: "false",
    bSearchable_3: "true",
    bSortable_3: "true",

    mDataProp_4: "4",
    sSearch_4: "",
    bRegex_4: "false",
    bSearchable_4: "true",
    bSortable_4: "true",

    mDataProp_5: "5",
    sSearch_5: "",
    bRegex_5: "false",
    bSearchable_5: "true",
    bSortable_5: "true",

    mDataProp_6: "6",
    sSearch_6: "",
    bRegex_6: "false",
    bSearchable_6: "true",
    bSortable_6: "true",

    mDataProp_7: "7",
    sSearch_7: "",
    bRegex_7: "false",
    bSearchable_7: "true",
    bSortable_7: "true",

    sSearch: "",
    bRegex: "false",
    iSortCol_0: "0",
    sSortDir_0: "asc",
    iSortingCols: "1",
    _: String(Date.now()),
  };

  const qs = Object.entries(params)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join("&");

  return `${ENDPOINT_PENYEDIA}?${qs}`;
}

function buildUrlSwakelola({ sEcho, start, length }) {
  const params = {
    idKldi: ID_KLDI,
    tahun: TAHUN,
    sEcho: String(sEcho),
    iColumns: "8",
    sColumns: ",,nama,,,,,",
    iDisplayStart: String(start),
    iDisplayLength: String(length),

    mDataProp_0: "0",
    sSearch_0: "",
    bRegex_0: "false",
    bSearchable_0: "true",
    bSortable_0: "false",

    mDataProp_1: "1",
    sSearch_1: "",
    bRegex_1: "false",
    bSearchable_1: "true",
    bSortable_1: "true",

    mDataProp_2: "2",
    sSearch_2: "",
    bRegex_2: "false",
    bSearchable_2: "true",
    bSortable_2: "true",

    mDataProp_3: "3",
    sSearch_3: "",
    bRegex_3: "false",
    bSearchable_3: "true",
    bSortable_3: "true",

    mDataProp_4: "4",
    sSearch_4: "",
    bRegex_4: "false",
    bSearchable_4: "true",
    bSortable_4: "true",

    mDataProp_5: "5",
    sSearch_5: "",
    bRegex_5: "false",
    bSearchable_5: "true",
    bSortable_5: "true",

    mDataProp_6: "6",
    sSearch_6: "",
    bRegex_6: "false",
    bSearchable_6: "true",
    bSortable_6: "true",

    mDataProp_7: "7",
    sSearch_7: "",
    bRegex_7: "false",
    bSearchable_7: "true",
    bSortable_7: "true",

    sSearch: "",
    bRegex: "false",
    iSortCol_0: "0",
    sSortDir_0: "asc",
    iSortingCols: "1",
    sRangeSeparator: "~",
    _: String(Date.now()),
  };

  const qs = Object.entries(params)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join("&");

  return `${ENDPOINT_SWAKELOLA}?${qs}`;
}

// =====================
// FETCH & WRITE GENERIC
// =====================
async function fetchDatatableAll({
  api,
  sheets,
  sheetName,
  referer,
  buildUrl,
  mapRow,
  headerRow,
  label,
}) {
  await ensureSheet(sheets, sheetName);
  await writeHeader(sheets, sheetName, headerRow);

  let start = 0;
  let sEcho = 1;
  let total = null;

  let buffer = [];
  let written = 0;

  for (let loop = 0; loop < 200000; loop++) {
    const url = buildUrl({ sEcho, start, length: PAGE_SIZE });

    const res = await api.get(url, {
      headers: {
        Accept: "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest",
        Referer: referer,
      },
    });

    const status = res.status();
    const text = await res.text();

    if (status !== 200) {
      console.log(`[${label}] HTTP`, status);
      console.log(`[${label}] Body head:`, text.slice(0, 500));
      throw new Error(`[${label}] Request ditolak / bukan 200`);
    }

    const head = text.trim().slice(0, 20).toLowerCase();
    if (head.startsWith("<!doctype") || head.startsWith("<html")) {
      console.log(`[${label}] DIBALES HTML (anti-bot). Head:`);
      console.log(text.slice(0, 500));
      throw new Error(`[${label}] Response HTML, bukan JSON`);
    }

    const json = JSON.parse(text);
    if (total == null) total = Number(json.iTotalRecords || 0);

    const aaData = json.aaData || [];
    if (!aaData.length) break;

    for (const row of aaData) buffer.push(mapRow(row));

    if (buffer.length >= FLUSH_EVERY) {
      await appendRows(sheets, sheetName, buffer);
      written += buffer.length;
      console.log(`[${label}] Written: ${written} / ${total || "?"}`);
      buffer = [];
    }

    start += PAGE_SIZE;
    sEcho += 1;

    if (total && start >= total) break;

    await sleep(SLEEP_MS);
  }

  if (buffer.length) {
    await appendRows(sheets, sheetName, buffer);
    written += buffer.length;
    console.log(`[${label}] Written final: ${written} / ${total || "?"}`);
  }

  console.log(`[${label}] DONE. Total rows:`, written);
}

// =====================
// MAIN
// =====================
async function main() {
  const sheets = await getSheetsClient();

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    locale: "id-ID",
    timezoneId: "Asia/Jakarta",
    userAgent:
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
  });

  const page = await context.newPage();

  // Buka 1 halaman dulu untuk cookie domain
  await page.goto(REFERER_PENYEDIA, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(800);

  const api = context.request;

  // === PENYEDIA ===
  await fetchDatatableAll({
    api,
    sheets,
    sheetName: SHEET_PENYEDIA,
    referer: REFERER_PENYEDIA,
    buildUrl: buildUrlPenyedia,
    headerRow: [
      "ID/No",
      "Satuan Kerja",
      "Nama Paket",
      "Pagu",
      "Metode",
      "Sumber Dana",
      "Kode RUP",
      "Waktu",
    ],
    mapRow: (row) => [
      row[0] ?? "",
      row[1] ?? "",
      row[2] ?? "",
      toNumber(row[3]),
      row[4] ?? "",
      row[5] ?? "",
      row[6] ?? "",
      row[7] ?? "",
    ],
    label: "PENYEDIA",
  });

  // Pastikan cookie masih kebawa (buka referer swakelola juga)
  await page.goto(REFERER_SWAKELOLA, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(800);

  // === SWAKELOLA ===
  await fetchDatatableAll({
    api,
    sheets,
    sheetName: SHEET_SWAKELOLA,
    referer: REFERER_SWAKELOLA,
    buildUrl: buildUrlSwakelola,
    headerRow: [
      "ID/No",
      "Satuan Kerja",
      "Nama",
      "Pagu",
      "Jenis Swakelola",
      "Sumber Dana",
      "Kode RUP",
      "Waktu",
    ],
    mapRow: (row) => [
      row[0] ?? "",
      row[1] ?? "",
      row[2] ?? "",
      toNumber(row[3]),
      row[4] ?? "",
      row[5] ?? "",
      row[6] ?? "",
      row[7] ?? "",
    ],
    label: "SWAKELOLA",
  });

  await browser.close();
  console.log("ALL DONE ✅");
}

main().catch((e) => {
  console.error("ERROR:", e);
  process.exit(1);
});
