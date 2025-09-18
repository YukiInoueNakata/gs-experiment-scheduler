/** ========= 基本ユーティリティ ========= */
function getSS_() {
  if (SS_ID) return SpreadsheetApp.openById(SS_ID);
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('スプレッドシートに紐づいていません。SS_ID を設定してください。');
  return ss;
}

/** ========= シート定義 ========= */
const SHEETS = {
  SLOTS: 'Slots',
  RESP: 'Responses',
  CONF: 'Confirmed',
  ARCH: 'Archive',
  AS: 'AddSlots',
  CO: 'CancelOps',
  MQ: 'MailQueue'
};

function getSlotHeaders() {
  return ['SlotID','Date','Start','End','Capacity','Location','Status','ConfirmedCount','Timezone'];
}

function getResponseHeaders() {
  return ['Timestamp','Name','Email','SlotID','Date','Start','End','Status','NotifiedConfirm','NotifiedWait','NotifiedRemind','Notes'];
}

function getConfirmedHeaders() {
  const base = ['SlotID','Date','Start','End','Location','ConfirmedAt'];
  for (let i = 1; i <= CONFIG.capacity; i++) {
    base.push(`Subject${i}Name`, `Subject${i}Email`);
  }
  base.push('ActualCount');
  return base;
}

function getArchiveHeaders() {
  return ['ArchivedAt','Timestamp','Name','Email','SlotID','Date','Start','End','Status','Notes','NotifiedConfirm','NotifiedWait','NotifiedRemind','RestoredAt'];
}

const MQ_HEADERS = ['CreatedAt','Type','To','Subject','Body','ICSText','MetaJson','Status','LastTriedAt','Error'];

/** ========= 初期化＆各シート ========= */
function ensureSheets_() {
  var ss = getSS_();
  if (!ss.getSheetByName(SHEETS.SLOTS)) {
    const sh = ss.insertSheet(SHEETS.SLOTS);
    sh.appendRow(getSlotHeaders());
  }
  if (!ss.getSheetByName(SHEETS.RESP)) {
    const sh = ss.insertSheet(SHEETS.RESP);
    sh.appendRow(getResponseHeaders());
  }
  ensureConfirmedSheet_();
  ensureArchiveSheet_();
  ensureMailQueueSheet_();
  ensureAddSlotsSheet_();
  ensureCancelOpsSheet_();
  removeDefaultSheet_();
}

function removeDefaultSheet_(){
  var ss = getSS_();
  ['シート1','Sheet1'].forEach(function(n){
    var sh = ss.getSheetByName(n);
    if (sh && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}

function ensureConfirmedSheet_(){
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.CONF);
  if (!sh) { 
    sh = ss.insertSheet(SHEETS.CONF); 
    sh.appendRow(getConfirmedHeaders()); 
  }
  return sh;
}

function ensureArchiveSheet_(){
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.ARCH);
  if (!sh) { 
    sh = ss.insertSheet(SHEETS.ARCH); 
    sh.appendRow(getArchiveHeaders()); 
  }
  return sh;
}

function ensureMailQueueSheet_(){
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.MQ);
  if (!sh) { 
    sh = ss.insertSheet(SHEETS.MQ); 
    sh.appendRow(MQ_HEADERS); 
  }
  return sh;
}

function ensureAddSlotsSheet_(){
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.AS);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.AS);
    sh.appendRow(['Mode','Date','Start','End','FromDate','ToDate','TimeWindows','ExcludeWeekends','Capacity','Location','Timezone','RespectConfigExcludes','Status','Result']);
    sh.appendRow(['date','2025-09-10','','','','','','FALSE',CONFIG.capacity,CONFIG.location,CONFIG.tz,'TRUE','example','← この行は見本です']);
    sh.setFrozenRows(1);
    var ruleMode = SpreadsheetApp.newDataValidation().requireValueInList(['datetime','date','range'], true).setAllowInvalid(false).build();
    sh.getRange('A2:A1000').setDataValidation(ruleMode);
    var ruleBool = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE','FALSE'], true).setAllowInvalid(false).build();
    sh.getRange('H2:H1000').setDataValidation(ruleBool);
    sh.getRange('L2:L1000').setDataValidation(ruleBool);
    sh.setColumnWidths(1, 12, 140);
  }
  return sh;
}

function ensureCancelOpsSheet_(){
  var ss = getSS_(); 
  var sh = ss.getSheetByName(SHEETS.CO);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.CO);
    sh.appendRow(['Email','Scope','SlotPolicy','FillPolicy','Reason','Status','Result']);
    sh.appendRow(['user@example.com','confirmed','refill-slot','try-fill','本人都合','example','← この行は見本です']);
    sh.setFrozenRows(1);
    var ruleScope = SpreadsheetApp.newDataValidation().requireValueInList(['confirmed','all'], true).setAllowInvalid(false).build();
    sh.getRange('B2:B1000').setDataValidation(ruleScope);
    var rulePolicy = SpreadsheetApp.newDataValidation().requireValueInList(['drop-slot','refill-slot'], true).setAllowInvalid(false).build();
    sh.getRange('C2:C1000').setDataValidation(rulePolicy);
    var ruleFill = SpreadsheetApp.newDataValidation().requireValueInList(['try-fill','keep-partial','to-pending','cancel-all'], true).setAllowInvalid(false).build();
    sh.getRange('D2:D1000').setDataValidation(ruleFill);
    sh.setColumnWidths(1, 7, 160);
  }
  return sh;
}

/** ========= 枠生成 ========= */
function setup() {
  ensureSheets_();
  generateSlotsFromConfig_();
  setupTriggers();
}

function clearSlots_(){
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  sh.clear(); 
  sh.appendRow(getSlotHeaders());
}

function generateSlotsFromConfig_(){
  clearSlots_();
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  var start = new Date(CONFIG.startDate+'T00:00:00'), end = new Date(CONFIG.endDate+'T00:00:00');
  var isExcludedDate = function(s){ return (CONFIG.excludeDates||[]).indexOf(s)>=0; };
  var isExcludedDT = function(d,st,en){ return (CONFIG.excludeDateTimes||[]).indexOf(d+' '+st+'-'+en)>=0; };
  
  for (var d=new Date(start); d<=end; d=new Date(d.getTime()+86400000)){
    if (CONFIG.excludeWeekends && (d.getDay()===0 || d.getDay()===6)) continue;
    var y=d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), da=('0'+d.getDate()).slice(-2);
    var dateStr = y+'-'+m+'-'+da;
    if (isExcludedDate(dateStr)) continue;
    CONFIG.timeWindows.forEach(function(win){
      var p=win.split('-'); 
      var st=p[0], en=p[1];
      if (isExcludedDT(dateStr, st, en)) return;
      createSlotRowIfNotExists_(dateStr, st, en, CONFIG.capacity, CONFIG.location, CONFIG.tz);
    });
  }
}

function createSlotRowIfNotExists_(dateStr, st, en, cap, loc, tz){
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  var id = dateStr + '_' + st.replace(':','');
  var vals = sh.getDataRange().getValues();
  for (var i=1;i<vals.length;i++){ 
    if (vals[i][0]===id) return false; 
  }
  sh.appendRow([id, dateStr, st, en, cap, loc, 'open', 0, tz]);
  return true;
}

/** ========= Webアプリ ========= */
function doGet() {
  var t = HtmlService.createTemplateFromFile('Index');
  t.title = CONFIG.title;
  t.consentHtml = TEMPLATES.consentHtml;
  t.capacity = CONFIG.capacity;
  return t.evaluate().setTitle(CONFIG.title);
}

function include(filename){ 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

function getSlots() {
  var sh = getSS_().getSheetByName(SHEETS.SLOTS);
  var values = sh.getDataRange().getValues(); 
  var head = values.shift();
  var resp = getResponses_(), bySlot = groupBy_(resp, function(r){ return r.SlotID; });

  var tomorrowStr = null;
  if (CONFIG.showOnlyFromTomorrow) {
    var n=new Date(), t=new Date(n.getFullYear(), n.getMonth(), n.getDate()+1);
    var y=t.getFullYear(), m=('0'+(t.getMonth()+1)).slice(-2), d=('0'+t.getDate()).slice(-2);
    tomorrowStr = y+'-'+m+'-'+d;
  }
  
  var out = values.map(function(row){
    var rec = asObj_(head,row);
    var ds = normDateStr_(rec.Date), st=normTimeStr_(rec.Start), en=normTimeStr_(rec.End);
    var slotResponses = bySlot[rec.SlotID]||[];
    var confirmed = slotResponses.filter(function(r){ return r.Status==='confirmed'; }).length;
    var pending = slotResponses.filter(function(r){ return r.Status==='pending'; }).length;
    var waitlist = slotResponses.filter(function(r){ return r.Status==='waitlist'; }).length;
    
    var label = (function(){ 
      var w='日月火水木金土'[ new Date(ds+'T00:00:00+09:00').getDay() ]; 
      return ds+' ('+w+')'; 
    })();
    
    // あと何名で確定かを計算
    var neededForConfirm = 0;
    var confirmMessage = '';
    
    if (confirmed >= CONFIG.minCapacityToConfirm) {
      // すでに最小人数を満たしている
      confirmMessage = '確定済み';
    } else if (pending + confirmed >= CONFIG.minCapacityToConfirm) {
      // pendingを含めれば最小人数を満たす
      neededForConfirm = 0;
      confirmMessage = '処理待ち';
    } else {
      // まだ最小人数に達していない
      neededForConfirm = CONFIG.minCapacityToConfirm - pending - confirmed;
      confirmMessage = 'あと' + neededForConfirm + '名で確定';
    }
    
    // 満席かどうかの情報も追加
    var isFull = confirmed >= Number(rec.Capacity);
    var availableSeats = Math.max(0, Number(rec.Capacity) - confirmed);
    
    return { 
      slotId:rec.SlotID, 
      date:ds, 
      dateLabel:label, 
      start:st, 
      end:en, 
      capacity:Number(rec.Capacity),
      status:rec.Status, 
      remaining:availableSeats,  // 空き席数
      confirmed:confirmed,        // 確定人数
      pending:pending,            // 申込み中の人数
      waitlist:waitlist,          // キャンセル待ちの人数
      neededForConfirm:neededForConfirm,  // あと何名で確定か
      confirmMessage:confirmMessage,       // 状態メッセージ
      isFull:isFull,              // 満席かどうか
      minCapacity:CONFIG.minCapacityToConfirm,  // 最小確定人数
      tz:rec.Timezone 
    };
  }).filter(function(s){ 
    return !tomorrowStr || s.date >= tomorrowStr; 
  }).sort(function(a,b){ 
    return (a.date+a.start).localeCompare(b.date+b.start); 
  });

  return { 
    title: CONFIG.title, 
    slots: out, 
    capacity: CONFIG.capacity,
    minCapacityToConfirm: CONFIG.minCapacityToConfirm  // フロントエンドでも使えるように追加
  };
}

/** ========= 申込処理 ========= */
function register(name, email, slotIds) {
  // ログ記録: 関数開始
  logFunctionStart_('register', {
    name: name,
    email: email,
    slotIds: slotIds,
    slotCount: slotIds ? slotIds.length : 0
  });

  // ユーザーアクション記録
  logUserAction_('申し込み開始', {
    name: name,
    email: email,
    requestedSlots: slotIds
  });

  try {
    // 入力値検証
    if (!name || !email || !slotIds || !slotIds.length) {
      const error = new Error('入力が不足しています。');
      logError_('register', error, { name, email, slotIds });
      throw error;
    }

    email = String(email).trim().toLowerCase();
    writeLog_('INFO', 'register', '入力値検証完了', {
      normalizedEmail: email,
      slotIds: slotIds
    });

    var lock = LockService.getScriptLock();
    writeLog_('DEBUG', 'register', 'スクリプトロック取得試行中');
    lock.waitLock(30000);
    writeLog_('DEBUG', 'register', 'スクリプトロック取得成功');

    try {
      var now = new Date();
      var ss = getSS_();
      var respSh = ss.getSheetByName(SHEETS.RESP);
      var slotSh = ss.getSheetByName(SHEETS.SLOTS);
      var slotsAll = readSheetAsObjects_(slotSh);

      writeLog_('INFO', 'register', 'スプレッドシート情報取得完了', {
        totalSlots: slotsAll.length,
        timestamp: now
      });

      // 既存の申し込み確認
      var existing = getResponses_().filter(function(r){
        return String(r.Email).toLowerCase()===email;
      });
      var already = new Set(existing.map(function(r){ return r.SlotID; }));

      writeLog_('INFO', 'register', '既存申し込み確認完了', {
        existingCount: existing.length,
        existingSlots: Array.from(already)
      });

      var created = [];
      var skipped = [];

      slotIds.forEach(function(id) {
        if (already.has(id)) {
          skipped.push({ slotId: id, reason: '既に申し込み済み' });
          return;
        }

        var slot = slotsAll.find(function(s){ return s.SlotID === id; });
        if (!slot) {
          skipped.push({ slotId: id, reason: 'スロットが見つからない' });
          return;
        }

        try {
          respSh.appendRow([now, name, email, id, slot.Date, slot.Start, slot.End, 'pending', false, false, false, '']);
          created.push({slotId:id, date:slot.Date, start:slot.Start, end:slot.End});

          writeLog_('INFO', 'register', '申し込み記録完了', {
            slotId: id,
            date: slot.Date,
            start: slot.Start,
            end: slot.End
          });
        } catch (error) {
          logError_('register', error, { slotId: id, slot: slot });
          skipped.push({ slotId: id, reason: 'データ記録エラー' });
        }
      });

      writeLog_('INFO', 'register', '申し込み処理完了', {
        createdCount: created.length,
        skippedCount: skipped.length,
        created: created,
        skipped: skipped
      });

      // 受付メール送信
      if (created.length) {
        try {
          var lines = created.map(function(s){
            var ds=normDateStr_(s.date), st=normTimeStr_(s.start), en=normTimeStr_(s.end);
            return '・'+fmtJPDateTime_(ds,st)+' - '+en+'（'+CONFIG.tz+'）';
          }).join('\n');

          var subject = renderTemplate_(TEMPLATES.participant.receiptSubject, {});
          var body = renderTemplate_(TEMPLATES.participant.receiptBody, {
            name: name,
            lines: lines,
            fromName: CONFIG.mailFromName
          });

          MailApp.sendEmail(email, subject, body, {name: CONFIG.mailFromName});

          writeLog_('INFO', 'register', '受付メール送信完了', {
            to: email,
            subject: subject,
            createdCount: created.length
          });
        } catch (error) {
          logError_('register', error, { email, created });
          // メール送信エラーは処理を止めない
        }
      }

      // バッチ処理スケジュール
      try {
        scheduleDelayedBatch_(CONFIG.batchProcessDelaySeconds || 30);
        writeLog_('INFO', 'register', 'バッチ処理スケジュール完了', {
          delaySeconds: CONFIG.batchProcessDelaySeconds || 30
        });
      } catch (error) {
        logError_('register', error, { delaySeconds: CONFIG.batchProcessDelaySeconds });
      }

      const result = {
        ok: true,
        message: '受付しました。確定の可否はメールでお知らせします。',
        created: created.length
      };

      logFunctionEnd_('register', result);
      logUserAction_('申し込み完了', {
        email: email,
        createdCount: created.length,
        result: result
      });

      return result;

    } finally {
      lock.releaseLock();
      writeLog_('DEBUG', 'register', 'スクリプトロック解放完了');
    }
  } catch (error) {
    logError_('register', error, { name, email, slotIds });
    logUserAction_('申し込みエラー', {
      email: email,
      error: error.message
    });
    throw error;
  }
}