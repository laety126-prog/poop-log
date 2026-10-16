/**
 * うんちきろく — 受け取り＆ノート用シート作成
 * Apps Script エディタのコードをこれに丸ごと差しかえて使います。
 */
const REC = '記録';
const NOTE = 'ノート用';
const CIRCLE = ['', '①', '②', '③', '④', '⑤', '⑥', '⑦'];
const SYM = {P: '痛がった', B: '血がついた', E: 'もれた', '浣': '浣腸', '座': '座薬'};
const HOURS = 24;
const FIRST_HOUR_COL = 3;               // C列 = 0時
const NOTE_COL = FIRST_HOUR_COL + HOURS; // AA列 = 気づいたことなど
const DAYS = 14;                        // ノート1ページ = 月曜始まりの2週間

/* ---------- アプリからの記録を受け取る ---------- */
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const t = String(p.time || '');
    const date = p.date || t.slice(0, 10);
    const hour = (p.hour !== undefined && p.hour !== '') ? Number(p.hour) : Number((t.slice(11).split(':')[0]) || 0);
    const symptoms = Array.isArray(p.symptoms) ? p.symptoms.join(',') : String(p.symptoms || '');
    recSheet_().appendRow([t, date, hour, String(p.amount || ''), Number(p.hardness) || '', symptoms, String(p.memo || ''), new Date()]);
    try { buildNote(); } catch (err) { console.error(err); }
    return json_({status: 'success'});
  } catch (error) {
    return json_({status: 'error', message: String(error)});
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function recSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(REC);
  if (!sh) {
    sh = ss.insertSheet(REC);
    sh.getRange('A:B').setNumberFormat('@');
    sh.appendRow(['日時', '日付', '時', '量', 'かたさ', '症状', 'メモ', '記録した日時']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const m = String(v).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  return m ? `${m[1]}-${('0' + m[2]).slice(-2)}-${('0' + m[3]).slice(-2)}` : '';
}

/* ---------- ノート用シートを作る ---------- */
function buildNote() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let note = ss.getSheetByName(NOTE);
  if (!note) { note = ss.insertSheet(NOTE, 0); }

  // 開始日（月曜）。B1のプルダウンで選んだ週。空なら「先週の月曜」
  const picked = note.getRange('B1').getValue();
  let start = null;
  if (picked instanceof Date) start = picked;
  else { const m = String(picked).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/); if (m) start = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])); }
  if (!start) { start = new Date(); start.setDate(start.getDate() - 7); }
  start = monday_(start);

  // 記録を日付×時間にまとめる
  const cells = {}, notes = {};
  const rows = recSheet_().getDataRange().getValues().slice(1);
  let earliest = null;
  for (const r of rows) {
    const d = dateKey_(r[1] || r[0]);
    if (!d) continue;
    if (!earliest || d < earliest) earliest = d;
    const h = Number(r[2]) || 0;
    const codes = String(r[5] || '').split(',').map(s => s.trim()).filter(Boolean);
    const text = `${r[3] || ''}\n${CIRCLE[Number(r[4])] || ''}${codes.length ? ' ' + codes.join(' ') : ''}`;
    (cells[`${d}|${h}`] = cells[`${d}|${h}`] || []).push(text);
    const words = codes.map(c => SYM[c] || c);
    if (r[6] && r[6] !== '旧データから変換') words.push(String(r[6]));
    if (words.length) (notes[d] = notes[d] || []).push(`${h}時 ${words.join('・')}`);
  }

  // 値をつくる
  const WD = ['日', '月', '火', '水', '木', '金', '土'];
  const header = ['月/日', '曜'];
  for (let h = 0; h < HOURS; h++) header.push(String(h));
  header.push('気づいたことなど');
  const values = [header];
  for (let i = 0; i < DAYS; i++) {
    const d = new Date(start); d.setDate(start.getDate() + i);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const row = [`${d.getMonth() + 1}/${d.getDate()}`, WD[d.getDay()]];
    for (let h = 0; h < HOURS; h++) row.push((cells[`${key}|${h}`] || []).join('\n'));
    row.push((notes[key] || []).join('／'));
    values.push(row);
  }

  // 書き込みと見た目
  const width = NOTE_COL;
  if (note.getMaxColumns() < width + 2) note.insertColumnsAfter(note.getMaxColumns(), width + 2 - note.getMaxColumns());
  note.getRange(1, 2, 1, width).breakApart();
  note.getRange('C1').clearContent();
  note.getRange(2, 1, 40, width + 2).clear();
  // 週を選ぶプルダウン（新しい週が上）
  const now = monday_(new Date());
  let first = earliest ? monday_(new Date(Number(earliest.slice(0, 4)), Number(earliest.slice(5, 7)) - 1, Number(earliest.slice(8, 10)))) : new Date(now);
  if (first > start) first = new Date(start);
  const labels = [];
  for (let w = new Date(now.getFullYear(), now.getMonth(), now.getDate()); w >= first && labels.length < 400; w.setDate(w.getDate() - 7)) labels.push(weekLabel_(w));
  const current = weekLabel_(start);
  if (labels.indexOf(current) < 0) labels.push(current);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(labels, true).setAllowInvalid(false).build();
  note.getRange('A1').setValue('表示する週');
  note.getRange('B1').setDataValidation(rule).setNumberFormat('@').setValue(current).setBackground('#fff7d6').setFontWeight('bold').setHorizontalAlignment('left');
  note.getRange('A1:B1').setFontSize(11);

  const all = note.getRange(2, 1, values.length, width);
  all.setValues(values)
     .setFontSize(11).setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true)
     .setBorder(true, true, true, true, true, true, '#c9b8b0', SpreadsheetApp.BorderStyle.SOLID);
  note.getRange(2, 1, 1, width).setBackground('#f6dfe6').setFontWeight('bold').setFontSize(10);
  note.getRange(3, NOTE_COL, DAYS, 1).setHorizontalAlignment('left').setFontSize(10);
  // 6時間ごとに太い線（ノートの0・6・12・18・24の目盛り）
  for (let h = 0; h < HOURS; h += 6) {
    note.getRange(2, FIRST_HOUR_COL + h, values.length, 6).setBorder(true, true, true, true, null, null, '#6d4c41', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  // 土日の色
  for (let i = 0; i < DAYS; i++) {
    const wd = values[i + 1][1];
    if (wd === '土') note.getRange(3 + i, 2).setFontColor('#1f6fd1').setFontWeight('bold');
    if (wd === '日') note.getRange(3 + i, 2).setFontColor('#d1342b').setFontWeight('bold');
  }
  note.setColumnWidth(1, 76); note.setColumnWidth(2, 32);
  note.getRange('B1:Z1').merge();
  note.setColumnWidths(FIRST_HOUR_COL, HOURS, 36);
  note.setColumnWidth(NOTE_COL, 280);
  note.setRowHeights(3, DAYS, 46);
  note.setFrozenRows(2); note.setFrozenColumns(2);
}

function monday_(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  x.setDate(x.getDate() - ((x.getDay() + 6) % 7));
  return x;
}
function weekLabel_(mon) {
  const end = new Date(mon); end.setDate(mon.getDate() + 13);
  return `${mon.getFullYear()}/${mon.getMonth() + 1}/${mon.getDate()}（月）〜${end.getMonth() + 1}/${end.getDate()}（日）の2週間`;
}

/* B1（表示する週）を選んだら作り直す */
function onEdit(e) {
  if (e && e.range.getSheet().getName() === NOTE && e.range.getA1Notation() === 'B1') buildNote();
}

/* メニュー */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('うんちきろく')
    .addItem('ノート用を更新', 'buildNote')
    .addItem('前の記録を取り込む（最初の1回だけ）', 'migrateOld')
    .addToUi();
}

/* ---------- 前のアプリの記録を変換して取り込む（1回だけ） ---------- */
function migrateOld() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('migrated')) { ss.toast('前の記録はもう取り込みずみです', 'うんちきろく'); return; }
  const old = ss.getSheets().find(s => s.getName() !== REC && s.getName() !== NOTE);
  if (!old) { ss.toast('前の記録のシートが見つかりませんでした', 'うんちきろく'); return; }
  const HARD = {'かたい': 2, 'ふつう': 4, 'やわらかい': 5};
  const AMT = {'すくない': 'S', 'ふつう': 'M', 'おおい': 'L'};
  const tz = Session.getScriptTimeZone();
  const out = [];
  for (const r of old.getDataRange().getValues()) {
    const tv = r[0];
    if (!tv || tv === '日時') continue;
    let date, hour;
    if (tv instanceof Date) { date = Utilities.formatDate(tv, tz, 'yyyy-MM-dd'); hour = tv.getHours(); }
    else { date = dateKey_(tv); const m = String(tv).match(/\s(\d{1,2}):/); hour = m ? Number(m[1]) : 0; }
    if (!date) continue;
    const hv = String(r[1]).trim(), av = String(r[2]).trim();
    const hard = HARD[hv] || (Number(hv) >= 1 && Number(hv) <= 7 ? Number(hv) : '');
    const amt = AMT[av] || (/^[SML]$/.test(av) ? av : '');
    out.push([`${date} ${hour}:00`, date, hour, amt, hard, '', '旧データから変換', new Date()]);
  }
  const rec = recSheet_();
  if (out.length) rec.getRange(rec.getLastRow() + 1, 1, out.length, 8).setValues(out);
  props.setProperty('migrated', '1');
  buildNote();
  ss.toast(`${out.length}件の記録を取り込みました`, 'うんちきろく');
}
