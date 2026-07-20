/**
 * MISI BENGKEL LEGO — UJIAN 1 DMD 3143
 * Backend: Google Apps Script
 * - Rekod semua jawapan pelajar ke Google Sheet (untuk kajian tindakan)
 * - Jana dokumen rasmi (PDF) daripada Google Docs template, diisi dengan jawapan pelajar
 *
 * CARA SETUP (rujuk arahan penuh yang disertakan bersama):
 * 1. Buat Google Sheet baru, buka Extensions > Apps Script, padam kod default,
 *    paste KESELURUHAN fail ini.
 * 2. Tukar TEMPLATE_DOC_ID di bawah dengan ID Google Docs template (selepas
 *    UJIAN1_DMD3143_Template_GoogleDocs.docx diupload & dibuka sebagai Google Docs).
 * 3. Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 * 4. Salin URL yang diberi, masukkan dalam index.html (APPS_SCRIPT_URL).
 */

const TEMPLATE_DOC_ID = "GANTI_DENGAN_ID_GOOGLE_DOCS_TEMPLATE"; // <-- WAJIB TUKAR
const SHEET_NAME = "Jawapan";
const OUTPUT_FOLDER_NAME = "UJIAN1_DMD3143_Jawapan_PDF"; // folder dalam Drive untuk simpan salinan PDF

const FIELD_ORDER = [
  "nama", "angka",
  "q1a_1", "q1a_2", "q1a_3", "q1b",
  "q2_225", "q2_50", "q2_r", "q2_16",
  "qb1", "qb2", "qb3_clip", "qb3_lekat", "qb4", "qb5a", "qb5b",
  "totalPoints", "answeredCount"
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    recordToSheet(data);
    const pdfBase64 = generateFilledDocument(data);
    return jsonOutput({ status: "ok", pdfBase64: pdfBase64 });
  } catch (err) {
    return jsonOutput({ status: "error", message: err.message });
  }
}

function doGet(e) {
  return jsonOutput({ status: "ok", message: "Misi Bengkel LEGO backend sedia. Guna POST untuk hantar jawapan." });
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp", "Nama", "Angka Giliran",
      "Q1a(i)", "Q1a(ii)", "Q1a(iii)", "Q1b",
      "Q2-225", "Q2-50", "Q2-R", "Q2-16",
      "QB1 (Prosedur Tyre Changer)", "QB2 (Keselamatan Balancing)",
      "QB3 Clip-on", "QB3 Lekat", "QB4", "QB5a", "QB5b",
      "Mata Gamifikasi", "Bil Soalan Dijawab"
    ]);
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function recordToSheet(data) {
  const sheet = getSheet();
  const a = data.answers || {};
  sheet.appendRow([
    new Date(),
    data.nama || "",
    data.angka || "",
    a.q1a_1 || "", a.q1a_2 || "", a.q1a_3 || "", a.q1b || "",
    a.q2_225 || "", a.q2_50 || "", a.q2_r || "", a.q2_16 || "",
    a.qb1 || "", a.qb2 || "",
    a.qb3_clip || "", a.qb3_lekat || "", a.qb4 || "", a.qb5a || "", a.qb5b || "",
    data.totalPoints || 0, data.answeredCount || 0
  ]);
}

function generateFilledDocument(data) {
  const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
  const folder = getOrCreateFolder(OUTPUT_FOLDER_NAME);
  const copyName = "UJIAN1_DMD3143_" + (data.nama || "pelajar").replace(/[^a-zA-Z0-9]+/g, "_") + "_" + new Date().getTime();
  const copyFile = templateFile.makeCopy(copyName, folder);
  const doc = DocumentApp.openById(copyFile.getId());
  const body = doc.getBody();

  const a = data.answers || {};
  const tarikh = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  const replacements = {
    "{{NAMA}}": data.nama || "-",
    "{{ANGKA}}": data.angka || "-",
    "{{TARIKH}}": tarikh,
    "{{TOTAL_POINTS}}": String(data.totalPoints || 0),
    "{{Q1A_1}}": a.q1a_1 || "-",
    "{{Q1A_2}}": a.q1a_2 || "-",
    "{{Q1A_3}}": a.q1a_3 || "-",
    "{{Q1B}}": a.q1b || "-",
    "{{Q2_225}}": a.q2_225 || "-",
    "{{Q2_50}}": a.q2_50 || "-",
    "{{Q2_R}}": a.q2_r || "-",
    "{{Q2_16}}": a.q2_16 || "-",
    "{{QB1}}": a.qb1 || "-",
    "{{QB2}}": a.qb2 || "-",
    "{{QB3_CLIP}}": a.qb3_clip || "-",
    "{{QB3_LEKAT}}": a.qb3_lekat || "-",
    "{{QB4}}": a.qb4 || "-",
    "{{QB5A}}": a.qb5a || "-",
    "{{QB5B}}": a.qb5b || "-"
  };

  Object.keys(replacements).forEach(function (key) {
    body.replaceText(escapeRegex(key), replacements[key]);
  });

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(copyFile.getId()).getAs(MimeType.PDF);
  return Utilities.base64Encode(pdfBlob.getBytes());
}

function escapeRegex(s) {
  return s.replace(/[{}]/g, "\\$&");
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

/** Jalankan fungsi ni SEKALI dalam editor Apps Script untuk authorize akses Drive/Docs */
function testAuthorize() {
  Logger.log(DriveApp.getRootFolder().getName());
  Logger.log(DocumentApp.openById(TEMPLATE_DOC_ID).getName());
}
