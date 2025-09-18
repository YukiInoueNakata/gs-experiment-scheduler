/** ========= 正規化 ========= */
function normDateStr_(v, zone) {
  var tz = zone || CONFIG.tz || 'Asia/Tokyo';
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  var d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  throw new Error('Invalid date: ' + s);
}

function normTimeStr_(v, zone) {
  var tz = zone || CONFIG.tz || 'Asia/Tokyo';
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm');
  var s = String(v || '').trim();
  if (/^\d{1,2}:\d{2}$/.test(s)) return s;
  if (/^\d{4}$/.test(s)) return s.slice(0,2)+':'+s.slice(2);
  var d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, tz, 'HH:mm');
  throw new Error('Invalid time: ' + s);
}

/** ========= 共通ヘルパ ========= */
function getResponses_(){ 
  var sh=getSS_().getSheetByName(SHEETS.RESP), vals=sh.getDataRange().getValues(), head=vals.shift(); 
  return vals.map(function(r){ return asObj_(head,r); }); 
}

function readSheetAsObjects_(sh){ 
  var vals=sh.getDataRange().getValues(), head=vals.shift(); 
  return vals.map(function(r){ return asObj_(head,r); }); 
}

function asObj_(head,row){ 
  var o={}; 
  head.forEach(function(h,i){ o[h]=row[i]; }); 
  return o; 
}

function groupBy_(arr,keyFn){ 
  return arr.reduce(function(m,x){ 
    var k=keyFn(x); 
    (m[k]||(m[k]=[])).push(x); 
    return m; 
  },{}); 
}

function colIndex_(head){ 
  var o={}; 
  head.forEach(function(h,i){ o[h]=i; }); 
  return o; 
}

function markNotified_(rowIndex, colName, val) {
  var sh=getSS_().getSheetByName(SHEETS.RESP), head=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var idx=head.indexOf(colName)+1; 
  sh.getRange(rowIndex, idx).setValue(val);
}

function markNotifiedByFind_(rec, colName, val) {
  const sh = getSS_().getSheetByName(SHEETS.RESP);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return false;
  const head = data[0], idx = colIndex_(head);
  const recTime = rec.Timestamp instanceof Date ? rec.Timestamp.getTime() : new Date(rec.Timestamp).getTime();
  const recEmail = String(rec.Email).toLowerCase();
  const recSlot  = String(rec.SlotID);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowTime = row[idx.Timestamp] instanceof Date ? row[idx.Timestamp].getTime() : new Date(row[idx.Timestamp]).getTime();
    const rowEmail = String(row[idx.Email]).toLowerCase();
    const rowSlot  = String(row[idx.SlotID]);
    if (rowTime === recTime && rowEmail === recEmail && rowSlot === recSlot) {
      row[idx[colName]] = !!val;
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return true;
    }
  }
  return false;
}

function deleteResponseRow_(rec){
  var sh=getSS_().getSheetByName(SHEETS.RESP), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  for (var i=0;i<vals.length;i++){
    var r=asObj_(head, vals[i]);
    if (r.Timestamp==rec.Timestamp && r.Email==rec.Email && r.SlotID==rec.SlotID){ 
      sh.deleteRow(i+2); 
      return true; 
    }
  } 
  return false;
}

function setResponseStatus_(rec, status){
  var sh=getSS_().getSheetByName(SHEETS.RESP), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head);
  for (var i=0;i<vals.length;i++){
    var r=asObj_(head, vals[i]);
    if (r.Timestamp==rec.Timestamp && r.Email==rec.Email && r.SlotID==rec.SlotID){
      vals[i][idx.Status]=status; 
      sh.getRange(i+2,1,1,vals[i].length).setValues([vals[i]]); 
      return true;
    }
  } 
  return false;
}

function hasConfirmedElsewhere_(email, excludeSlotId) {
  const confirmed = getResponses_().filter(r =>
    String(r.Email).toLowerCase() === email.toLowerCase() &&
    r.Status === 'confirmed' &&
    r.SlotID !== excludeSlotId
  );
  return confirmed.length > 0;
}

/** ========= ログ機能 ========= */

/**
 * 詳細ログを記録するための関数
 * @param {string} level - ログレベル (INFO, WARN, ERROR, DEBUG)
 * @param {string} function_name - 実行中の関数名
 * @param {string} message - ログメッセージ
 * @param {Object} data - 追加データ (オプション)
 */
function writeLog_(level, function_name, message, data = {}) {
  try {
    const timestamp = new Date();
    const logEntry = {
      timestamp: timestamp,
      level: level,
      function: function_name,
      message: message,
      data: JSON.stringify(data),
      user: Session.getActiveUser().getEmail() || 'unknown'
    };

    // コンソールログも出力
    console.log(`[${level}] ${function_name}: ${message}`, data);

    // ログシートに記録
    const logSheet = ensureLogSheet_();
    logSheet.appendRow([
      timestamp,
      level,
      function_name,
      message,
      JSON.stringify(data),
      logEntry.user
    ]);

  } catch (error) {
    // ログ記録でエラーが発生した場合はコンソールのみに出力
    console.error('Failed to write log:', error);
    console.log(`[${level}] ${function_name}: ${message}`, data);
  }
}

/**
 * ログ専用シートを作成・取得
 */
function ensureLogSheet_() {
  const ss = getSS_();
  let logSheet = ss.getSheetByName('SystemLog');

  if (!logSheet) {
    logSheet = ss.insertSheet('SystemLog');
    logSheet.appendRow([
      'Timestamp', 'Level', 'Function', 'Message', 'Data', 'User'
    ]);

    // ヘッダーを固定
    logSheet.setFrozenRows(1);

    // 列幅を調整
    logSheet.setColumnWidths(1, 6, [150, 80, 150, 300, 200, 150]);

    // ヘッダーの背景色を設定
    logSheet.getRange(1, 1, 1, 6).setBackground('#f0f0f0').setFontWeight('bold');
  }

  return logSheet;
}

/**
 * エラー情報を詳細に記録
 * @param {string} function_name - エラーが発生した関数名
 * @param {Error} error - エラーオブジェクト
 * @param {Object} context - エラー発生時の文脈情報
 */
function logError_(function_name, error, context = {}) {
  const errorData = {
    name: error.name,
    message: error.message,
    stack: error.stack,
    context: context
  };

  writeLog_('ERROR', function_name, `エラーが発生しました: ${error.message}`, errorData);
}

/**
 * 関数の開始をログに記録
 * @param {string} function_name - 関数名
 * @param {Object} params - 関数のパラメータ
 */
function logFunctionStart_(function_name, params = {}) {
  writeLog_('DEBUG', function_name, '関数開始', { parameters: params });
}

/**
 * 関数の終了をログに記録
 * @param {string} function_name - 関数名
 * @param {Object} result - 関数の実行結果
 */
function logFunctionEnd_(function_name, result = {}) {
  writeLog_('DEBUG', function_name, '関数終了', { result: result });
}

/**
 * ユーザーアクションをログに記録
 * @param {string} action - アクション名
 * @param {Object} details - アクションの詳細
 */
function logUserAction_(action, details = {}) {
  writeLog_('INFO', 'USER_ACTION', `ユーザーアクション: ${action}`, details);
}

/**
 * バッチ処理の詳細をログに記録
 * @param {string} step - 処理ステップ
 * @param {Object} details - 処理の詳細
 */
function logBatchProcess_(step, details = {}) {
  writeLog_('INFO', 'BATCH_PROCESS', `バッチ処理: ${step}`, details);
}