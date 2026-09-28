/*==============================================================
  【機材在庫管理システム】スタジオエレファント
  Code.gs  —  v2（原価・上限裁定・検証付き同期）
  --------------------------------------------------------------
  初回のみ setup() を実行。既に稼働中の場合は upgrade() を実行。
==============================================================*/

const MASTER_ID  = '1np_LDoEpsavp-uwAGNvwwxTTTFs_li4VzrFE1Ay62P4';
const MASTER_TABS = ['映像', '音響', '照明', 'その他'];

const P = PropertiesService.getScriptProperties();
const SH_ITEMS = '機材マスター';
const SH_LOG   = '入出荷ログ';
const SH_EST   = '見積データ';
const SH_CFG   = '設定';
const SH_RATE  = '掛け率';
const COLS = 14;   // A〜N（Nは原価）

const ITEM_HEADER = ['ID','カテゴリ','品目','機材名','数量','公開備考','社内備考',
                     'URL','単価','料金タイプ','貸出中','状態','更新日時','原価'];

const DEFAULT_RATES = [[0,0.5],[1,1],[2,1.3],[3,1.5],[4,1.7],[5,1.9],[6,2.15],[7,2.4],[8,2.65],
  [9,2.9],[10,3.15],[11,3.4],[12,3.65],[13,3.9],[14,4.15],[15,4.4],[16,4.65],[17,4.9],[18,5.15],
  [19,5.4],[20,5.65],[21,5.9],[22,6.15],[23,6.4],[24,6.65],[25,6.9],[26,7.15],[27,7.4],[28,7.65],
  [29,7.9],[30,8.15],[31,8.4],
  [32,8.65],[33,8.9],[34,9.15],[35,9.4],[36,9.65],[37,9.9],[38,10.15],[39,10.4],[40,10.65],
  [41,10.9],[42,11.15],[43,11.4],[44,11.65],[45,11.9],[46,12.15],[47,12.4],[48,12.65],[49,12.9],
  [50,13.15],[51,13.4],[52,13.65],[53,13.9],[54,14.15],[55,14.4],[56,14.65],[57,14.9],[58,15.15],
  [59,15.4],[60,15.65],[61,15.9],[62,16.15]];

const DEFAULT_CFG = {
  company:'有限会社スタジオエレファント',
  address:'〒130-0011 東京都墨田区石原3-2-3 両国鈴木ビル1F/2F',
  tel:'03-5839-2942', fax:'03-5839-2943',
  tax:'10', threshold:'2', pageSize:'50',
  rateCap:'0',      // Aタイプ掛率の上限（0＝無制限）
  costRatio:'0',    // 1点あたり金額が「原価×この％」を超えたら警告（0＝無効）
  admins:''
};

/*------------------ セットアップ / アップグレード ------------------*/
function setup() {
  let id = P.getProperty('APP_SS_ID');
  let ss;
  if (id) { ss = SpreadsheetApp.openById(id); }
  else {
    ss = SpreadsheetApp.create('機材在庫管理_アプリデータ');
    P.setProperty('APP_SS_ID', ss.getId());
    ss.getSheets()[0].setName(SH_ITEMS);
  }
  ensureSheet_(ss, SH_ITEMS, ITEM_HEADER);
  ensureSheet_(ss, SH_LOG,  ['日時','種別','機材ID','機材名','数量','案件/備考','操作者']);
  ensureSheet_(ss, SH_EST,  ['見積ID','見積コード','案件名','担当者','場所','支払条件','本番日',
                             '搬入日','搬出日','備考','明細JSON','小計','合計','作成日','作成者',
                             '更新日時','更新者','ロック']);
  ensureSheet_(ss, SH_CFG,  ['キー','値']);
  ensureSheet_(ss, SH_RATE, ['日数','掛率']);
  ensureCols_(ss);
  ensureEstCols_(ss);
  seedCfg_(ss);
  const has = ScriptApp.getProjectTriggers().some(function(t){
    return t.getHandlerFunction() === 'syncFromMaster'; });
  if (!has) ScriptApp.newTrigger('syncFromMaster').timeBased().atHour(6).everyDays(1).create();

  const msg = 'セットアップ完了\nアプリデータ: ' + ss.getUrl() +
    '\n同期: 追加' + r.added + ' / 更新' + r.updated + ' / 廃番' + r.retired;
  Logger.log(msg); return msg;
}

/** 稼働中のシートにv2の列・設定を追加する（データは消えません） */
function upgrade() {
  const ss = appSS_();
  ensureCols_(ss);
  ensureEstCols_(ss);
  seedCfg_(ss);
  const msg = 'アップグレード完了：原価列・見積更新履歴列・ロック列・上限設定を追加しました';
  Logger.log(msg); return msg;
}

function ensureSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.getRange(1,1,1,header.length).setFontWeight('bold')
      .setBackground('#1B3A5B').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function ensureCols_(ss) {
  const sh = ss.getSheetByName(SH_ITEMS);
  if (sh.getLastColumn() < COLS) {
    sh.getRange(1, COLS).setValue('原価')
      .setFontWeight('bold').setBackground('#1B3A5B').setFontColor('#ffffff');
  }
}
/** 既存の「見積データ」シートに更新日時・更新者・ロックの列が無ければ追加する（データは消えません） */
function ensureEstCols_(ss) {
  const sh = ss.getSheetByName(SH_EST);
  if (!sh) return;
  if (sh.getLastColumn() < 17) {
    sh.getRange(1,16,1,2).setValues([['更新日時','更新者']])
      .setFontWeight('bold').setBackground('#1B3A5B').setFontColor('#ffffff');
  }
  if (sh.getLastColumn() < 18) {
    sh.getRange(1,18).setValue('ロック')
      .setFontWeight('bold').setBackground('#1B3A5B').setFontColor('#ffffff');
  }
}
function seedCfg_(ss) {
  const sh = ss.getSheetByName(SH_CFG);
  const have = {};
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2,1,last-1,2).getValues().forEach(function(r){ if(r[0]) have[r[0]]=true; });
  const add = Object.keys(DEFAULT_CFG).filter(function(k){ return !have[k]; })
                    .map(function(k){ return [k, DEFAULT_CFG[k]]; });
  if (add.length) sh.getRange(sh.getLastRow()+1,1,add.length,2).setValues(add);

  const rs = ss.getSheetByName(SH_RATE);
  if (rs.getLastRow() < 2) rs.getRange(2,1,DEFAULT_RATES.length,2).setValues(DEFAULT_RATES);
}
function appSS_() {
  const id = P.getProperty('APP_SS_ID');
  if (!id) throw new Error('未セットアップです。エディタで setup() を実行してください。');
  return SpreadsheetApp.openById(id);
}

/*------------------ 画面 ------------------*/
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action) {
    var args = (e.parameter.args) ? JSON.parse(e.parameter.args) : [];
    return handleApi_(action, args);
  }
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('機材在庫管理システム｜スタジオエレファント')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*------------------ 外部フロントエンド（GitHub Pages）向けAPI ------------------
   画面(index.html)をGAS外（GitHub Pages）でホストするようになったため、
   従来のgoogle.script.runの代わりに、このWebアプリURLへのfetch()で呼び出す。
   POSTは Content-Type: text/plain のJSON文字列で {action, args} を受け取り、
   対応する既存関数をそのまま呼び出して結果をJSONで返す。 -------------------- */
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return handleApi_(body.action, body.args || []);
}
var API_ACTIONS_ = {
  getBootstrap: getBootstrap,
  manualSync: manualSync,
  savePricing: savePricing,
  applyMoves: applyMoves,
  returnByEstimate: returnByEstimate,
  updateEstimate: updateEstimate,
  saveEstimate: saveEstimate,
  duplicateEstimate: duplicateEstimate,
  deleteEstimate: deleteEstimate,
  setEstimateLock: setEstimateLock,
  saveRates: saveRates,
  saveConfig: saveConfig,
  exportCsv: exportCsv
};
function handleApi_(action, args) {
  var out;
  try {
    var fn = API_ACTIONS_[action];
    if (!fn) throw new Error('未知のaction: ' + action);
    out = JSON.stringify(fn.apply(null, args || []));
  } catch (err) {
    out = JSON.stringify({ __error: String((err && err.message) || err) });
  }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function currentUser_() {
  try { return Session.getActiveUser().getEmail() || '(不明)'; } catch(e){ return '(不明)'; }
}

/*------------------ 診断 ------------------*/
function findSheet_(ss, name) {
  const norm = function(s){ return String(s).replace(/[\s　]/g,''); };
  const t = norm(name);
  const hit = ss.getSheets().filter(function(sh){ return norm(sh.getName()) === t; });
  return hit.length ? hit[0] : null;
}
function checkMaster() {
  const master = SpreadsheetApp.openById(MASTER_ID);
  const out = ['■ ' + master.getName(), '■ ' + master.getUrl(), '', '■ シート一覧'];
  master.getSheets().forEach(function(sh){
    out.push('  「'+sh.getName()+'」 行:'+sh.getLastRow()+' 列:'+sh.getLastColumn()); });
  out.push('');
  MASTER_TABS.forEach(function(tab){
    const sh = findSheet_(master, tab);
    out.push('─────────────');
    if (!sh) { out.push('× タブ「'+tab+'」が見つかりません'); return; }
    out.push('○ 「'+sh.getName()+'」 行:'+sh.getLastRow());
    if (sh.getLastRow() < 1) return;
    const h = sh.getRange(1,1,1,Math.min(8,sh.getLastColumn())).getValues()[0];
    out.push('  見出し: ' + h.join(' | '));
    const a1 = String(h[0]).trim();
    if (a1 !== 'ID') { out.push('  ★NG: A1が「'+a1+'」です'); return; }
    const a2 = sh.getLastRow()>=2 ? String(sh.getRange(2,1).getValue()).trim() : '';
    out.push(/^[VALO]-\d{4}$/i.test(a2) ? '  → 形式OK' : '  ★NG: A2が「'+a2+'」です');
  });
  const m = out.join('\n'); Logger.log(m); return m;
}

/*------------------ 同期（検証つき） ------------------*/
function syncFromMaster() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const master = SpreadsheetApp.openById(MASTER_ID);
    const incoming = {}; const errors = []; const skipped = [];
    let read = 0;

    MASTER_TABS.forEach(function(tab){
      const sh = findSheet_(master, tab);
      if (!sh) { errors.push('タブ「'+tab+'」が見つかりません'); return; }
      if (sh.getLastRow() < 2) { errors.push('タブ「'+tab+'」にデータがありません'); return; }
      if (String(sh.getRange(1,1).getValue()).trim() !== 'ID') {
        errors.push('タブ「'+tab+'」のA1が「ID」ではないため取り込みを中止しました'); return;
      }
      sh.getRange(2,1,sh.getLastRow()-1,7).getValues().forEach(function(r,i){
        const id = String(r[0]).trim().toUpperCase();
        if (!id) return;
        if (!/^[VALO]-\d{4}$/.test(id)) { skipped.push(tab+' '+(i+2)+'行: ID不正「'+id+'」'); return; }
        if (incoming[id]) { skipped.push(tab+' '+(i+2)+'行: ID重複「'+id+'」'); return; }
        const q = Number(r[3]);
        if (isNaN(q)) { skipped.push(tab+' '+(i+2)+'行: 数量が数値でない'); return; }
        read++;
        incoming[id] = { id:id, cat:tab, hinmoku:String(r[1]).trim(), name:String(r[2]).trim(),
          total:q, pub:String(r[4]||''), naibu:String(r[5]||''), url:String(r[6]||'') };
      });
    });

    if (read === 0) return { ok:false, added:0, updated:0, retired:0, read:0,
      warnings:[], errors: errors.length?errors:['1件も読み込めませんでした'], skipped:skipped.slice(0,20) };

    const sh = appSS_().getSheetByName(SH_ITEMS);
    const last = sh.getLastRow();
    const cur = last > 1 ? sh.getRange(2,1,last-1,COLS).getValues() : [];
    const now = new Date();
    let added=0, updated=0, retired=0; const warnings=[];

    cur.forEach(function(r,i){
      const id = String(r[0]).trim().toUpperCase();
      if (!id) return;
      const inc = incoming[id];
      if (!inc) { if (r[11] !== '廃番') { cur[i][11]='廃番'; cur[i][12]=now; retired++; } return; }
      const rented = Number(r[10])||0;
      if (inc.total < rented) warnings.push(inc.name+'：数量'+inc.total+'台に対し貸出中'+rented+'台');
      const ch = (r[1]!==inc.cat||r[2]!==inc.hinmoku||r[3]!==inc.name||Number(r[4])!==inc.total||
                  r[5]!==inc.pub||r[6]!==inc.naibu||r[7]!==inc.url||r[11]!=='有効');
      if (ch) {
        cur[i][1]=inc.cat; cur[i][2]=inc.hinmoku; cur[i][3]=inc.name; cur[i][4]=inc.total;
        cur[i][5]=inc.pub; cur[i][6]=inc.naibu; cur[i][7]=inc.url;
        cur[i][11]='有効'; cur[i][12]=now; updated++;
      }
      delete incoming[id];
    });

    const news = Object.keys(incoming).sort().map(function(id){
      const x = incoming[id]; added++;
      return [x.id,x.cat,x.hinmoku,x.name,x.total,x.pub,x.naibu,x.url,0,'A',0,'有効',now,0];
    });
    if (cur.length)  sh.getRange(2,1,cur.length,COLS).setValues(cur);
    if (news.length) sh.getRange(cur.length+2,1,news.length,COLS).setValues(news);

    return { ok:true, added:added, updated:updated, retired:retired, read:read,
             warnings:warnings, errors:errors, skipped:skipped.slice(0,20) };
  } finally { lock.releaseLock(); }
}
function manualSync() {
  const r = syncFromMaster();
  if (r.ok) r.items = readItems_(appSS_());
  return r;
}
function resetItems() {
  const ss = appSS_();
  const sh = ss.getSheetByName(SH_ITEMS);
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,COLS).clearContent();
  const r = syncFromMaster();
  const m = r.ok ? ('取り込み直しました：'+r.added+'件（読込'+r.read+'行）')
                 : ('中止：\n'+(r.errors||[]).join('\n'));
  Logger.log(m); return m;
}

/*------------------ 読み込み ------------------*/
function getBootstrap() {
  const ss = appSS_();
  let masterUrl='', masterError='';
  try { masterUrl = SpreadsheetApp.openById(MASTER_ID).getUrl(); }
  catch(e){ masterError = '公開マスターにアクセスできません。閲覧権限の共有が必要です。'; }
  return { user:currentUser_(), isAdmin:isAdmin_(), items:readItems_(ss),
    history:readLog_(ss,200), estimates:readEstimates_(ss), cfg:readCfg_(ss),
    rates:readRates_(ss), masterUrl:masterUrl, masterError:masterError, appUrl:ss.getUrl() };
}
function readItems_(ss) {
  const sh = ss.getSheetByName(SH_ITEMS);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,COLS).getValues().filter(function(r){return r[0];})
    .map(function(r){ return { id:String(r[0]), cat:r[1], hinmoku:r[2], name:r[3],
      total:Number(r[4])||0, pub:r[5], naibu:r[6], url:r[7], price:Number(r[8])||0,
      ptype:r[9]||'A', rented:Number(r[10])||0, status:r[11]||'有効', cost:Number(r[13])||0 }; });
}
function readLog_(ss, limit) {
  const sh = ss.getSheetByName(SH_LOG);
  const last = sh.getLastRow(); if (last < 2) return [];
  const st = Math.max(2, last-limit+1);
  return sh.getRange(st,1,last-st+1,7).getValues().reverse().map(function(r){
    return { date:fmt_(r[0]), type:r[1], itemId:r[2], name:r[3],
             qty:Number(r[4])||0, note:r[5], user:r[6] }; });
}
function readEstimates_(ss) {
  const sh = ss.getSheetByName(SH_EST);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,18).getValues().filter(function(r){return r[0];})
    .reverse().map(function(r){ return { id:r[0], code:r[1], project:r[2], person:r[3],
      place:r[4], pay:r[5], honban:d_(r[6]), hannyuu:d_(r[7]), hanshutsu:d_(r[8]), biko:r[9],
      lines:JSON.parse(r[10]||'[]'), subtotal:Number(r[11])||0, total:Number(r[12])||0,
      date:fmt_(r[13]), user:r[14], updated:fmt_(r[15]), updatedBy:r[16]||'', locked:r[17]||'' }; });
}
function readCfg_(ss) {
  const sh = ss.getSheetByName(SH_CFG);
  const o = {}; Object.keys(DEFAULT_CFG).forEach(function(k){ o[k]=DEFAULT_CFG[k]; });
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,2).getValues()
    .forEach(function(r){ if(r[0]) o[r[0]] = r[1]; });
  return o;
}
function readRates_(ss) {
  const sh = ss.getSheetByName(SH_RATE); const o = {};
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(function(r){
    if (r[0]==='' || r[0]===null) return; o[String(Number(r[0]))] = Number(r[1]); });
  return o;
}
function isAdmin_() {
  const list = String(readCfg_(appSS_()).admins||'').split(',')
    .map(function(s){return s.trim().toLowerCase();}).filter(Boolean);
  return !list.length || list.indexOf(currentUser_().toLowerCase()) >= 0;
}
function fmt_(v){ return !v ? '' : (v instanceof Date
  ? Utilities.formatDate(v,'Asia/Tokyo','yyyy/MM/dd HH:mm') : String(v)); }
function d_(v){ return !v ? '' : (v instanceof Date
  ? Utilities.formatDate(v,'Asia/Tokyo','yyyy-MM-dd') : String(v)); }

/*------------------ 在庫の増減 ------------------*/
function applyMoves(moves, note) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = appSS_(); const sh = ss.getSheetByName(SH_ITEMS);
    const data = sh.getRange(2,1,sh.getLastRow()-1,COLS).getValues();
    const idx = {}; data.forEach(function(r,i){ if(r[0]) idx[String(r[0])]=i; });
    const errs = [];
    moves.forEach(function(m){
      const i = idx[m.id];
      if (i === undefined) { errs.push(m.id+' が見つかりません'); return; }
      const total=Number(data[i][4])||0, rented=Number(data[i][10])||0;
      const next = rented + Number(m.delta);
      if (next < 0)     errs.push(data[i][3]+'：貸出中'+rented+'台に対し返却'+(-m.delta)+'台');
      if (next > total) errs.push(data[i][3]+'：在庫'+(total-rented)+'台に対し出荷'+m.delta+'台');
    });
    if (errs.length) return { ok:false, errors:errs };

    const now = new Date(), user = currentUser_(), logs = [];
    moves.forEach(function(m){
      const i = idx[m.id];
      data[i][10] = (Number(data[i][10])||0) + Number(m.delta);
      data[i][12] = now;
      logs.push([now, Number(m.delta)>0?'出荷':'入荷', data[i][0], data[i][3],
                 Math.abs(Number(m.delta)), note||'', user]);
    });
    sh.getRange(2,1,data.length,COLS).setValues(data);
    const ls = ss.getSheetByName(SH_LOG);
    if (logs.length) ls.getRange(ls.getLastRow()+1,1,logs.length,7).setValues(logs);
    return { ok:true, items:readItems_(ss), history:readLog_(ss,200) };
  } finally { lock.releaseLock(); }
}

/*------------------ 単価・原価 ------------------*/
function savePricing(list) {
  if (!isAdmin_()) return { ok:false, errors:['単価の変更は管理者のみ可能です'] };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = appSS_(); const sh = ss.getSheetByName(SH_ITEMS);
    const data = sh.getRange(2,1,sh.getLastRow()-1,COLS).getValues();
    const idx = {}; data.forEach(function(r,i){ if(r[0]) idx[String(r[0])]=i; });
    const now = new Date();
    list.forEach(function(x){
      const i = idx[x.id]; if (i === undefined) return;
      if (x.price !== undefined && x.price !== null) data[i][8]  = Number(x.price)||0;
      if (x.ptype) data[i][9] = x.ptype;
      if (x.cost  !== undefined && x.cost  !== null) data[i][13] = Number(x.cost)||0;
      data[i][12] = now;
    });
    sh.getRange(2,1,data.length,COLS).setValues(data);
    return { ok:true, items:readItems_(ss) };
  } finally { lock.releaseLock(); }
}

/*------------------ 見積 ------------------*/
function saveEstimate(est) {
  const ss = appSS_(); const sh = ss.getSheetByName(SH_EST); const now = new Date();
  const id = 'EST-' + Utilities.formatDate(now,'Asia/Tokyo','yyyyMMdd-HHmmss');
  sh.appendRow([id, est.code||'', est.project||'', est.person||'', est.place||'', est.pay||'',
    est.honban||'', est.hannyuu||'', est.hanshutsu||'', est.biko||'',
    JSON.stringify(est.lines||[]), est.subtotal||0, est.total||0, now, currentUser_()]);
  ss.getSheetByName(SH_LOG).appendRow([now,'見積','',est.project||'',0,
    '見積作成 '+(est.code||'')+' ¥'+Number(est.total||0).toLocaleString(), currentUser_()]);
  return { ok:true, id:id, estimates:readEstimates_(ss), history:readLog_(ss,200) };
}
/** 既存の見積を上書き更新する（新しい行は作らない。作成日・作成者は保持し、更新日時・更新者を別欄に記録） */
function updateEstimate(id, est) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = appSS_(); const sh = ss.getSheetByName(SH_EST);
    const last = sh.getLastRow();
    if (last < 2) return { ok:false, errors:['見積データがありません'] };
    const data = sh.getRange(2,1,last-1,18).getValues();
    let rowIdx = -1;
    for (let i=0;i<data.length;i++){ if (String(data[i][0])===id) { rowIdx=i; break; } }
    if (rowIdx < 0) return { ok:false, errors:['対象の見積が見つかりません（削除された可能性があります）'] };
    const lockedBy = data[rowIdx][17];
    const me = currentUser_();
    if (lockedBy && lockedBy !== me && !isAdmin_()) {
      return { ok:false, errors:['この見積は「'+lockedBy+'」がロック中のため編集できません。複製してから編集してください。'] };
    }
    const now = new Date();
    const origCreatedDate = data[rowIdx][13];   // 作成日（保持する）
    const origCreatedBy   = data[rowIdx][14];   // 作成者（保持する＝所有者）
    sh.getRange(rowIdx+2,1,1,18).setValues([[
      id, est.code||'', est.project||'', est.person||'', est.place||'', est.pay||'',
      est.honban||'', est.hannyuu||'', est.hanshutsu||'', est.biko||'',
      JSON.stringify(est.lines||[]), est.subtotal||0, est.total||0,
      origCreatedDate, origCreatedBy, now, me, lockedBy
    ]]);
    ss.getSheetByName(SH_LOG).appendRow([now,'見積','',est.project||'',0,
      '見積修正 '+(est.code||'')+' ¥'+Number(est.total||0).toLocaleString(), me]);
    return { ok:true, id:id, estimates:readEstimates_(ss), history:readLog_(ss,200) };
  } finally { lock.releaseLock(); }
}
/** 見積の所有者（作成者）または管理者のみ、ロックの掛け外しができる */
function setEstimateLock(id, wantLock) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = appSS_(); const sh = ss.getSheetByName(SH_EST);
    const last = sh.getLastRow();
    if (last < 2) return { ok:false, errors:['見積データがありません'] };
    const data = sh.getRange(2,1,last-1,18).getValues();
    let rowIdx = -1;
    for (let i=0;i<data.length;i++){ if (String(data[i][0])===id) { rowIdx=i; break; } }
    if (rowIdx < 0) return { ok:false, errors:['対象の見積が見つかりません'] };
    const owner = data[rowIdx][14], me = currentUser_();
    if (owner !== me && !isAdmin_()) {
      return { ok:false, errors:['ロックの変更は、所有者（'+owner+'）または管理者のみ行えます'] };
    }
    sh.getRange(rowIdx+2,18).setValue(wantLock ? me : '');
    return { ok:true, estimates:readEstimates_(ss) };
  } finally { lock.releaseLock(); }
}
/** 見積の所有者（作成者）または管理者のみ、削除できる（ロックの有無に関わらず） */
function deleteEstimate(id) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = appSS_(); const sh = ss.getSheetByName(SH_EST);
    const last = sh.getLastRow();
    if (last < 2) return { ok:false, errors:['見積データがありません'] };
    const data = sh.getRange(2,1,last-1,18).getValues();
    let rowIdx = -1;
    for (let i=0;i<data.length;i++){ if (String(data[i][0])===id) { rowIdx=i; break; } }
    if (rowIdx < 0) return { ok:false, errors:['対象の見積が見つかりません（すでに削除済みの可能性があります）'] };
    const owner = data[rowIdx][14], me = currentUser_(), project = data[rowIdx][2];
    if (owner !== me && !isAdmin_()) {
      return { ok:false, errors:['削除は、所有者（'+owner+'）または管理者のみ行えます'] };
    }
    sh.deleteRow(rowIdx+2);
    const now = new Date();
    ss.getSheetByName(SH_LOG).appendRow([now,'見積','',project||'',0,'見積削除',me]);
    return { ok:true, estimates:readEstimates_(ss), history:readLog_(ss,200) };
  } finally { lock.releaseLock(); }
}
/** 見積を複製する。ロック中でも常に可能。複製した本人が新しい所有者になる（ロックは引き継がない） */
function duplicateEstimate(id) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = appSS_(); const sh = ss.getSheetByName(SH_EST);
    const last = sh.getLastRow();
    if (last < 2) return { ok:false, errors:['見積データがありません'] };
    const data = sh.getRange(2,1,last-1,18).getValues();
    let src = null;
    for (let i=0;i<data.length;i++){ if (String(data[i][0])===id) { src=data[i]; break; } }
    if (!src) return { ok:false, errors:['複製元の見積が見つかりません'] };
    const now = new Date();
    const newId = 'EST-' + Utilities.formatDate(now,'Asia/Tokyo','yyyyMMdd-HHmmss');
    const me = currentUser_();
    sh.appendRow([newId, src[1], (src[2]||'') + '（複製）', src[3], src[4], src[5],
      src[6], src[7], src[8], src[9], src[10], src[11], src[12], now, me, '', '', '']);
    ss.getSheetByName(SH_LOG).appendRow([now,'見積','',(src[2]||'')+'（複製）',0,
      '見積複製 元:'+(src[1]||''), me]);
    return { ok:true, id:newId, estimates:readEstimates_(ss), history:readLog_(ss,200) };
  } finally { lock.releaseLock(); }
}
function returnByEstimate(estId) {
  const ss = appSS_(); const sh = ss.getSheetByName(SH_EST);
  const rows = sh.getRange(2,1,sh.getLastRow()-1,15).getValues();
  const hit = rows.filter(function(r){ return r[0]===estId; })[0];
  if (!hit) return { ok:false, errors:['見積が見つかりません'] };
  const lines = JSON.parse(hit[10]||'[]').filter(function(l){ return l.id; });
  if (!lines.length) return { ok:false, errors:['返却対象の機材がありません'] };
  return applyMoves(lines.map(function(l){ return { id:l.id, delta:-Number(l.qty) }; }),
                    '一括返却「'+hit[2]+'」');
}

/*------------------ 設定 ------------------*/
function saveConfig(cfg) {
  if (!isAdmin_()) return { ok:false, errors:['設定の変更は管理者のみ可能です'] };
  const ss = appSS_(); const sh = ss.getSheetByName(SH_CFG);
  const rows = Object.keys(cfg).map(function(k){ return [k, cfg[k]]; });
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,2).clearContent();
  if (rows.length) sh.getRange(2,1,rows.length,2).setValues(rows);
  return { ok:true, cfg:readCfg_(ss) };
}
function saveRates(rates) {
  if (!isAdmin_()) return { ok:false, errors:['掛け率の変更は管理者のみ可能です'] };
  const ss = appSS_(); const sh = ss.getSheetByName(SH_RATE);
  const rows = Object.keys(rates).map(function(k){ return [Number(k), Number(rates[k])]; })
    .sort(function(a,b){ return a[0]-b[0]; });
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,2).clearContent();
  if (rows.length) sh.getRange(2,1,rows.length,2).setValues(rows);
  return { ok:true, rates:readRates_(ss) };
}
function exportCsv(filename, csvText) {
  const fs = DriveApp.getFoldersByName('機材管理_出力');
  const folder = fs.hasNext() ? fs.next() : DriveApp.createFolder('機材管理_出力');
  return folder.createFile(Utilities.newBlob('\uFEFF'+csvText,'text/csv',filename)).getUrl();
}
