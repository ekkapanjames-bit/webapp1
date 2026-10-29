/**
 * ระบบสร้างเอกสารอัตโนมัติจากเทมเพลต (Code.gs)
 * อัปโหลดเทมเพลต -> ตรวจจับสัญลักษณ์ {{ชื่อ}} -> กรอกข้อมูลในเว็บ -> แทนที่ -> ได้ไฟล์ PDF/DOCX
 * ไม่ใช้ Google Sheet: เทมเพลตเก็บในโฟลเดอร์ย่อย Templates, ไฟล์ที่สร้างเก็บใน Output
 */
const folder_id = "1YyEjlORuztH0GWs_BpCCGIwMr4hAew1u"; // โฟลเดอร์หลักใน Google Drive
const TZ = "Asia/Bangkok";
const MAX_FIELDS = 100;
const MAX_UPLOAD_MB = 20;
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PAIRS = { "{{": "}}", "[[": "]]", "{": "}", "[": "]" }; // สัญลักษณ์ที่รองรับ

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("ระบบสร้างเอกสารอัตโนมัติ")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- Helper ----------
function sub_(name) {
  const root = DriveApp.getFolderById(folder_id);
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}
function safeName_(s) { return String(s || "").replace(/[\\\/:*?"<>|\r\n]+/g, "-").trim().substring(0, 120); }
function esc_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
function delim_(open, close) {
  return PAIRS[open] === close ? [open, close] : ["{{", "}}"];
}
function makeRe_(open, close, flags) {
  const d = delim_(open, close);
  const chars = esc_(Array.from(new Set((d[0] + d[1]).split(""))).join(""));
  return new RegExp(esc_(d[0]) + "\\s*([^" + chars + "\\n]{1,100}?)\\s*" + esc_(d[1]), flags);
}
// อนุญาตเฉพาะ Google Docs ที่อยู่ในโฟลเดอร์ Templates เท่านั้น
function assertTemplate_(id) {
  let f;
  try { f = DriveApp.getFileById(id); } catch (e) { throw new Error("ไม่พบเทมเพลต"); }
  const tid = sub_("Templates").getId();
  const ps = f.getParents();
  while (ps.hasNext()) {
    if (ps.next().getId() === tid && f.getMimeType() === MimeType.GOOGLE_DOCS && !f.isTrashed()) return f;
  }
  throw new Error("ไม่พบเทมเพลต");
}
function thaiDate_() {
  const mn = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  const n = new Date();
  return Number(Utilities.formatDate(n, TZ, "d")) + " " + mn[Number(Utilities.formatDate(n, TZ, "M")) - 1] + " " + (Number(Utilities.formatDate(n, TZ, "yyyy")) + 543);
}

// ---------- เทมเพลต ----------
function getBootData() {
  return { templates: listTemplates(), max: MAX_FIELDS, today: thaiDate_() };
}

function listTemplates() {
  const out = [];
  const it = sub_("Templates").getFilesByType(MimeType.GOOGLE_DOCS);
  while (it.hasNext()) { const f = it.next(); out.push({ id: f.getId(), name: f.getName() }); }
  return out.sort(function (a, b) { return a.name.localeCompare(b.name, "th"); });
}

// อัปโหลด .docx/.doc แล้วแปลงเป็น Google Docs (ต้องเปิด Drive API ใน appsscript.json)
function uploadTemplate(p) {
  if (!p || !p.data || !p.name) throw new Error("ไม่พบไฟล์");
  if (!/\.docx?$/i.test(p.name)) throw new Error("รองรับเฉพาะไฟล์ .docx หรือ .doc");
  if (p.data.length > MAX_UPLOAD_MB * 1.37 * 1024 * 1024) throw new Error("ไฟล์ใหญ่เกิน " + MAX_UPLOAD_MB + " MB");
  if (typeof Drive === "undefined") throw new Error("ยังไม่ได้เปิดบริการ Drive API (วางไฟล์ appsscript.json ตามคู่มือ)");
  const title = safeName_(p.name.replace(/\.docx?$/i, "")) || "เทมเพลต";
  const blob = Utilities.newBlob(Utilities.base64Decode(p.data), /\.doc$/i.test(p.name) ? "application/msword" : DOCX_MIME, p.name);
  const f = Drive.Files.create({ name: title, mimeType: MimeType.GOOGLE_DOCS, parents: [sub_("Templates").getId()] }, blob);
  return { id: f.id, name: title };
}

// เพิ่มเทมเพลตจากลิงก์ Google Docs (คัดลอกเข้าโฟลเดอร์ Templates ไม่แตะไฟล์ต้นฉบับ)
function addTemplateFromLink(url) {
  const m = String(url || "").match(/[-\w]{25,}/);
  if (!m) throw new Error("ลิงก์ไม่ถูกต้อง");
  let f;
  try { f = DriveApp.getFileById(m[0]); } catch (e) { throw new Error("เปิดไฟล์ไม่ได้ ตรวจสอบสิทธิ์การเข้าถึง"); }
  if (f.getMimeType() !== MimeType.GOOGLE_DOCS) throw new Error("ต้องเป็นไฟล์ Google Docs เท่านั้น");
  const c = f.makeCopy(f.getName(), sub_("Templates"));
  return { id: c.getId(), name: c.getName() };
}

function deleteTemplate(id) {
  assertTemplate_(id).setTrashed(true);
  return true;
}

// ---------- ตรวจจับสัญลักษณ์ ----------
function detectFields(templateId, open, close) {
  assertTemplate_(templateId);
  const doc = DocumentApp.openById(templateId);
  let text = doc.getBody().getText();
  if (doc.getHeader()) text += "\n" + doc.getHeader().getText();
  if (doc.getFooter()) text += "\n" + doc.getFooter().getText();
  const re = makeRe_(open, close, "g");
  const seen = {}, out = [];
  let m;
  while ((m = re.exec(text))) {
    const k = m[1].trim();
    if (k && !seen[k]) { seen[k] = 1; out.push(k); }
  }
  return { fields: out.slice(0, MAX_FIELDS), total: out.length };
}

// ---------- สร้างเอกสาร ----------
function exportFile_(id, fmt, fname) {
  const r = UrlFetchApp.fetch("https://docs.google.com/document/d/" + id + "/export?format=" + fmt, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) throw new Error("ส่งออกไฟล์ " + fmt + " ไม่สำเร็จ");
  return { name: fname, mime: fmt === "pdf" ? "application/pdf" : DOCX_MIME, data: Utilities.base64Encode(r.getContent()) };
}

function generateDocument(p) {
  const tpl = assertTemplate_(p.templateId);
  const vals = p.values || {};
  const keys = Object.keys(vals);
  if (keys.length > MAX_FIELDS) throw new Error("ใช้ช่องกรอกได้ไม่เกิน " + MAX_FIELDS + " ช่อง");
  const d = delim_(p.open, p.close);
  const name = safeName_(p.name) || safeName_(tpl.getName() + "_" + Utilities.formatDate(new Date(), TZ, "yyyyMMdd-HHmmss"));
  const copy = tpl.makeCopy(name, sub_("Output"));
  try {
    const doc = DocumentApp.openById(copy.getId());
    const parts = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(Boolean);
    keys.forEach(function (k) {
      const key = String(k).trim();
      if (!key || key.length > 100) return;
      const pat = esc_(d[0]) + "\\s*" + esc_(key) + "\\s*" + esc_(d[1]);
      const val = String(vals[k] == null ? "" : vals[k]).replace(/[\\$]/g, "\\$&");
      parts.forEach(function (x) { x.replaceText(pat, val); });
    });
    doc.saveAndClose();
    const out = {
      name: name,
      pdf: exportFile_(copy.getId(), "pdf", name + ".pdf"),
      docx: exportFile_(copy.getId(), "docx", name + ".docx"),
      docUrl: null
    };
    if (p.keep) out.docUrl = copy.getUrl(); else copy.setTrashed(true);
    return out;
  } catch (e) {
    copy.setTrashed(true);
    throw e;
  }
}
