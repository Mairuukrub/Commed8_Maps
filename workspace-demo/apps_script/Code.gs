/**
 * COMMED 69 เขต 8 · ศูนย์รวมวิถีชุมชน — ตัวรับข้อมูล (Google Apps Script, Web app)
 *
 * เก็บใน Google Sheet (สร้างให้เองตอนรัน setup) + รูปใน Google Drive (โฟลเดอร์ COMMED8-workspace-photos)
 * รหัสเข้าใช้อยู่ในชีต "access" (code | name | role | wi | note | last_login) เพิ่ม/ลบรหัสได้เองโดยไม่ต้อง Deploy ใหม่
 *   role: admin (ผู้ดูแล ทำได้ทุกอย่าง) · lead (หัวหน้าวิถี ใส่เลข 1-6 ในคอลัมน์ wi) · member (สมาชิก) · viewer (ดูอย่างเดียว)
 *
 * ทุกคำขอเป็น POST { action, code, ... } (ข้อความล้วน ไม่ติด CORS preflight) คืน JSON
 *   login        → { ok, name, role, wi }
 *   pull         { since } → ข้อมูลที่เปลี่ยนหลังเวลา since: { now, recs: { notes|snips|people|feats|rels: [{id, rev, at, by, json}] } }
 *   push         { changes:[{ t, id, base, data }] } → ผลรายการ ok / conflict (มีคนแก้ก่อน คืนฉบับเซิร์ฟเวอร์) / rejected / toolarge
 *   uploadPhoto  { mime, b64 } → { id }          photo { id } → { mime, b64 }   (เฉพาะไฟล์ในโฟลเดอร์รูปของระบบนี้)
 *   codes / addCode / revokeCode  (admin เท่านั้น)
 *   purge        { t: feats|people|rels, id, base, label } → ลบถาวรรายการที่ซ่อนอยู่แล้ว (admin เท่านั้น) · แถวเหลือแค่ id + purged
 *
 * ไม่มีรหัสใดอยู่ในโค้ดนี้ · ห้ามใส่ข้อมูลจริงใน repo สาธารณะ
 */
const TABS = ["notes", "snips", "people", "feats", "rels"];
const ACCESS = "access";
const HEAD = ["id", "rev", "at", "by", "json"];
const MAX_JSON = 48000;        // ขีดจำกัด 50,000 ตัวอักษรต่อเซลล์ของ Google Sheets
const MAX_B64 = 4500000;       // รูปหลังย่อแล้ว (ฝั่งเว็บย่อเหลือ ~100-300 KB) เผื่อไว้
const FOLDER_NAME = "COMMED8-workspace-photos";
const ROLES = ["admin", "lead", "member", "viewer"];
const STATUSES = ["new", "checked", "pick", "skip"];
const props_ = () => PropertiesService.getScriptProperties();

/** รันครั้งเดียวใน editor: สร้าง Sheet + โฟลเดอร์รูป + รหัสผู้ดูแลชุดแรก (ดูที่ชีต access) */
function setup() {
  const props = props_();
  let ss = props.getProperty("SHEET_ID") ? SpreadsheetApp.openById(props.getProperty("SHEET_ID")) : null;
  if (!ss) { ss = SpreadsheetApp.create("COMMED8 ศูนย์รวมวิถีชุมชน (ข้อมูล)"); props.setProperty("SHEET_ID", ss.getId()); }
  TABS.forEach((t) => {
    if (!ss.getSheetByName(t)) { const sh = ss.insertSheet(t); sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]); sh.setFrozenRows(1); }
  });
  if (!ss.getSheetByName(ACCESS)) {
    const ac = ss.insertSheet(ACCESS);
    ac.getRange(1, 1, 1, 6).setValues([["code", "name", "role", "wi", "note", "last_login"]]);
    ac.getRange("A:A").setNumberFormat("@");
    ac.setFrozenRows(1);
    ac.appendRow([newCode_(), "ผู้ดูแล", "admin", "", "เปลี่ยนชื่อได้ · role: admin / lead / member / viewer · lead ใส่เลขวิถี 1-6 ในคอลัมน์ wi", ""]);
  }
  const def = ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  if (!props.getProperty("FOLDER_ID")) props.setProperty("FOLDER_ID", DriveApp.createFolder(FOLDER_NAME).getId());
  if (!props.getProperty("LAST")) props.setProperty("LAST", "0");
  Logger.log("Sheet: " + ss.getUrl());
  return ss.getUrl();
}

const newCode_ = () => Utilities.getUuid().replace(/-/g, "").slice(0, 12);
let SS_ = null;
function ss_() {
  if (SS_) return SS_;
  const id = props_().getProperty("SHEET_ID");
  if (!id) throw new Error("ยังไม่ได้รัน setup");
  SS_ = SpreadsheetApp.openById(id);
  return SS_;
}
/** คืนแท็บของข้อมูล ถ้ายังไม่มี (เช่น แท็บ rels ที่เพิ่มทีหลัง) จะสร้างให้เองพร้อมหัวตาราง ไม่ต้องรัน setup() ซ้ำ */
const tab_ = (t) => {
  const ss = ss_();
  let sh = ss.getSheetByName(t);
  if (!sh && TABS.indexOf(t) >= 0) { sh = ss.insertSheet(t); sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]); sh.setFrozenRows(1); }
  return sh;
};
const json_ = (o) => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet() {
  return json_({ ok: true, app: "commed8-workspace", v: 1, ready: !!props_().getProperty("SHEET_ID") });
}

function auth_(code) {
  code = String(code || "").trim();
  if (code.length < 6) return null;
  const sh = tab_(ACCESS);
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === code && ROLES.indexOf(String(v[i][2]).trim()) >= 0)
      return { name: String(v[i][1] || "ไม่ระบุชื่อ"), role: String(v[i][2]).trim(), wi: Number(v[i][3]) || 0, row: i + 2 };
  }
  return null;
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "คำขอไม่ถูกต้อง" }); }
  try {
    const user = auth_(body.code);
    if (!user) return json_({ ok: false, error: "รหัสไม่ถูกต้อง", code: "auth" });
    switch (body.action) {
      case "login": try { tab_(ACCESS).getRange(user.row, 6).setValue(new Date().toISOString()); } catch (er) {} return json_({ ok: true, name: user.name, role: user.role, wi: user.wi });
      case "pull": return json_(pull_(body));
      case "push": return json_(push_(user, body));
      case "uploadPhoto": return json_(uploadPhoto_(user, body));
      case "photo": return json_(photo_(body));
      case "purge":
        if (user.role !== "admin") return json_({ ok: false, error: "เฉพาะผู้ดูแล" });
        return json_(purge_(user, body));
      case "codes": case "addCode": case "revokeCode":
        if (user.role !== "admin") return json_({ ok: false, error: "เฉพาะผู้ดูแล" });
        return json_(adminAction_(user, body));
      default: return json_({ ok: false, error: "ไม่รู้จักคำสั่ง" });
    }
  } catch (err) { return json_({ ok: false, error: "เซิร์ฟเวอร์ผิดพลาด: " + err.message }); }
}

// ---------------------------------------------------------------- อ่าน
function pull_(body) {
  const since = Number(body.since) || 0;
  const last = Number(props_().getProperty("LAST") || 0);
  if (since > 0 && since >= last) return { ok: true, now: last, recs: {} };   // ไม่มีอะไรใหม่ ไม่ต้องเปิดชีต
  const recs = {};
  TABS.forEach((t) => {
    const sh = tab_(t), n = sh.getLastRow() - 1;
    recs[t] = n < 1 ? [] : sh.getRange(2, 1, n, 5).getValues()
      .filter((r) => Number(r[2]) > since)
      .map((r) => ({ id: String(r[0]), rev: Number(r[1]), at: Number(r[2]), by: String(r[3]), json: String(r[4]) }));
  });
  return { ok: true, now: last, recs: recs };
}

// ---------------------------------------------------------------- เขียน (ล็อกกันชนกัน)
const canTriage_ = (user, w) => user.role === "admin" || (user.role === "lead" && user.wi === Number(w));

/** สถานะรายวิถีแก้ได้เฉพาะหัวหน้าวิถีนั้น/ผู้ดูแล · ข้อความแก้ได้เฉพาะเจ้าของหรือผู้ดูแล · ผู้เขียนเดิมห้ามเปลี่ยน */
function guardSnip_(user, old, d) {
  let mod = false;
  if (!Array.isArray(d.wi)) { d.wi = []; mod = true; }
  const oldSt = (old && old.st) || {}, nw = d.st && typeof d.st === "object" ? d.st : {}, st = {};
  Object.keys(Object.assign({}, oldSt, nw)).forEach((w) => {
    let v;
    if (canTriage_(user, w)) v = nw[w] !== undefined ? nw[w] : oldSt[w];
    else if (oldSt[w] !== undefined) { v = oldSt[w]; if (nw[w] !== oldSt[w]) mod = true; }
    else { v = "new"; if (nw[w] !== undefined && nw[w] !== "new") mod = true; }
    if (STATUSES.indexOf(v) < 0) { v = "new"; mod = true; }
    st[w] = v;
  });
  d.st = st;
  if (old) {
    if (d.by !== old.by) { d.by = old.by; mod = true; }
    if (d.text !== old.text && user.role !== "admin" && old.by !== user.name) { d.text = old.text; mod = true; }
  }
  return { data: d, mod: mod };
}

function push_(user, body) {
  if (user.role === "viewer") return { ok: false, error: "บทบาทผู้ดูอย่างเดียว (viewer) แก้ไขไม่ได้" };
  const ch = body.changes || [];
  if (!Array.isArray(ch) || ch.length > 60) return { ok: false, error: "ส่งครั้งละไม่เกิน 60 รายการ" };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const props = props_();
    let last = Number(props.getProperty("LAST") || 0);
    const ix = {};
    const index = (t) => {
      if (ix[t]) return ix[t];
      const m = {}, sh = tab_(t), n = sh.getLastRow() - 1;
      if (n > 0) sh.getRange(2, 1, n, 5).getValues().forEach((r, i) => { m[String(r[0])] = { row: i + 2, rev: Number(r[1]), json: String(r[4]) }; });
      return (ix[t] = m);
    };
    const results = [], appended = [];
    ch.forEach((c) => {
      const t = c && c.t, id = String((c && c.id) || "");
      if (TABS.indexOf(t) < 0 || !id || id.length > 40 || !c.data || typeof c.data !== "object" || Array.isArray(c.data)) { results.push({ t: t, id: id, status: "rejected", error: "รูปแบบไม่ถูกต้อง" }); return; }
      const m = index(t), old = m[id];
      if (old && old.row === 0) { results.push({ t: t, id: id, status: "rejected", error: "รหัสซ้ำในคำขอเดียวกัน" }); return; }
      if (old && Number(c.base) !== old.rev) { results.push({ t: t, id: id, status: "conflict", rev: old.rev, json: old.json }); return; }
      let data = c.data, mod = false;
      if (t === "snips") { const g = guardSnip_(user, old ? JSON.parse(old.json) : null, data); data = g.data; mod = g.mod; }
      data.id = id;
      const json = JSON.stringify(data);
      if (json.length > MAX_JSON) { results.push({ t: t, id: id, status: "toolarge", error: "ข้อมูลยาวเกินที่ Google Sheets เก็บได้ในหนึ่งเซลล์ ลองแยกเป็นหลายโน้ต" }); return; }
      last = Math.max(Date.now(), last + 1);
      const rev = (old ? old.rev : 0) + 1, row = [id, rev, last, user.name, json];
      if (old) tab_(t).getRange(old.row, 1, 1, 5).setValues([row]);
      else { appended.push([t, row]); m[id] = { row: 0, rev: rev, json: json }; }
      if (old) { old.rev = rev; old.json = json; }
      const r = { t: t, id: id, status: "ok", rev: rev, at: last };
      if (mod) r.json = json;
      results.push(r);
    });
    appended.forEach((a) => tab_(a[0]).appendRow(a[1]));
    props.setProperty("LAST", String(last));
    return { ok: true, now: last, results: results };
  } finally { lock.releaseLock(); }
}

// ---------------------------------------------------------------- ลบถาวร (เฉพาะผู้ดูแล · เฉพาะรายการที่ "ซ่อน" ไว้แล้ว)
// แถวไม่ถูกลบออกจากชีต แต่เนื้อหาถูกแทนด้วย {"id":..,"purged":1} (rev เพิ่ม) เพื่อให้เครื่องอื่นที่ดึงข้อมูลแบบเพิ่มทีละส่วนรู้ว่าต้องลบ
// รูปของหมุดถูกย้ายไปถังขยะของ Drive (กู้ได้ 30 วัน) · บันทึกว่าใครลบอะไรไว้ในแท็บ purge_log
const PURGEABLE = ["feats", "people", "rels"];
const PURGE_LOG = "purge_log";
function purge_(user, body) {
  const t = String(body.t || ""), id = String(body.id || "");
  if (PURGEABLE.indexOf(t) < 0 || !id || id.length > 40) return { ok: false, error: "รูปแบบไม่ถูกต้อง" };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = tab_(t), n = sh.getLastRow() - 1;
    let hit = -1, vals = [];
    if (n > 0) { vals = sh.getRange(2, 1, n, 5).getValues(); for (let i = 0; i < n; i++) if (String(vals[i][0]) === id) { hit = i; break; } }
    if (hit < 0) return { ok: false, error: "ไม่พบรายการนี้" };
    const r = vals[hit];
    let d = {};
    try { d = JSON.parse(String(r[4])); } catch (er) {}
    if (d && d.purged) return { ok: true, already: true, rev: Number(r[1]), at: Number(r[2]), now: Number(props_().getProperty("LAST") || 0), trashed: 0 };
    if (Number(body.base) !== Number(r[1])) return { ok: false, code: "stale", error: "มีคนแก้รายการนี้ก่อน ลองใหม่อีกครั้ง" };
    if (d.hidden !== true) return { ok: false, error: "ต้องกด “ซ่อน” รายการนี้ก่อน จึงลบถาวรได้" };
    if (d.mergedInto) return { ok: false, error: "รายการนี้ถูกรวมไว้ ต้องกดยกเลิกการรวมก่อน" };
    let trashed = 0;
    if (t === "feats" && Array.isArray(d.photos)) {
      d.photos.forEach((p) => {
        const m = /^drive:([\w-]{10,80})$/.exec(String(p));
        if (!m) return;
        try {
          const file = DriveApp.getFileById(m[1]), want = props_().getProperty("FOLDER_ID"), ps = file.getParents();
          let inside = false;
          while (ps.hasNext()) if (ps.next().getId() === want) inside = true;
          if (inside) { file.setTrashed(true); trashed++; }
        } catch (er) {}
      });
    }
    const props = props_();
    const last = Math.max(Date.now(), Number(props.getProperty("LAST") || 0) + 1);
    const rev = Number(r[1]) + 1;
    sh.getRange(hit + 2, 1, 1, 5).setValues([[id, rev, last, user.name, JSON.stringify({ id: id, purged: 1 })]]);
    props.setProperty("LAST", String(last));
    const ss = ss_();
    let lg = ss.getSheetByName(PURGE_LOG);
    if (!lg) { lg = ss.insertSheet(PURGE_LOG); lg.getRange(1, 1, 1, 6).setValues([["time", "by", "table", "id", "label", "photos_trashed"]]); lg.setFrozenRows(1); }
    lg.appendRow([new Date().toISOString(), user.name, t, id, String(body.label || "").slice(0, 80), trashed]);
    return { ok: true, rev: rev, at: last, now: last, trashed: trashed };
  } finally { lock.releaseLock(); }
}

// ---------------------------------------------------------------- รูป (Drive)
const folder_ = () => DriveApp.getFolderById(props_().getProperty("FOLDER_ID"));
function uploadPhoto_(user, body) {
  if (user.role === "viewer") return { ok: false, error: "บทบาทผู้ดูอย่างเดียว อัปโหลดรูปไม่ได้" };
  const mime = String(body.mime || ""), b64 = String(body.b64 || "");
  if (["image/jpeg", "image/png", "image/webp"].indexOf(mime) < 0) return { ok: false, error: "รับเฉพาะรูป jpeg/png/webp" };
  if (!b64 || b64.length > MAX_B64) return { ok: false, error: "รูปใหญ่เกินไป" };
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, "p-" + Date.now() + (mime === "image/png" ? ".png" : ".jpg"));
  return { ok: true, id: folder_().createFile(blob).getId() };
}
function photo_(body) {
  const id = String(body.id || "");
  if (!/^[\w-]{10,80}$/.test(id)) return { ok: false, error: "รหัสรูปไม่ถูกต้อง" };
  const file = DriveApp.getFileById(id), want = props_().getProperty("FOLDER_ID");
  let inside = false;
  const ps = file.getParents();
  while (ps.hasNext()) if (ps.next().getId() === want) inside = true;
  if (!inside) return { ok: false, error: "ไม่ใช่รูปของระบบนี้" };   // กันไม่ให้ดึงไฟล์อื่นใน Drive ผ่านรหัสทีม
  const blob = file.getBlob();
  return { ok: true, mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) };
}

// ---------------------------------------------------------------- จัดการรหัส (admin)
function adminAction_(user, body) {
  const sh = tab_(ACCESS);
  const rows = () => (sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues());
  if (body.action === "codes") return { ok: true, rows: rows().map((r) => ({ code: String(r[0]), name: String(r[1]), role: String(r[2]), wi: Number(r[3]) || 0, note: String(r[4]), last: String(r[5]) })) };
  if (body.action === "addCode") {
    const name = String(body.name || "").trim(), role = String(body.role || "");
    const wi = Number(body.wi) || 0;
    if (!name) return { ok: false, error: "ใส่ชื่อก่อน" };
    if (ROLES.indexOf(role) < 0) return { ok: false, error: "บทบาทไม่ถูกต้อง" };
    if (role === "lead" && !(wi >= 1 && wi <= 6)) return { ok: false, error: "หัวหน้าวิถีต้องระบุวิถี 1-6" };
    const code = newCode_();
    sh.appendRow([code, name, role, role === "lead" ? wi : "", String(body.note || ""), ""]);
    return { ok: true, code: code };
  }
  // revokeCode
  const target = String(body.target || "").trim();
  if (target === String(body.code).trim()) return { ok: false, error: "ลบรหัสของตัวเองไม่ได้" };
  const rs = rows();
  for (let i = 0; i < rs.length; i++) {
    if (String(rs[i][0]).trim() === target) {
      if (String(rs[i][2]).trim() === "admin" && rs.filter((r) => String(r[2]).trim() === "admin").length < 2) return { ok: false, error: "ต้องเหลือผู้ดูแลอย่างน้อยหนึ่งคน" };
      sh.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: "ไม่พบรหัส" };
}
