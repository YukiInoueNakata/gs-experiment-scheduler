/** ========= バッチ処理 ========= */
function scheduleDelayedBatch_(seconds) {
  logBatchProcess_('バッチ処理スケジュール', {
    delaySeconds: seconds,
    scheduledAt: new Date()
  });

  ScriptApp.newTrigger('processPendingBatch_')
    .timeBased()
    .after(seconds * 1000)
    .create();
}

function processPendingBatch_() {
  logBatchProcess_('バッチ処理開始', {
    startTime: new Date(),
    user: Session.getActiveUser().getEmail()
  });

  const lock = LockService.getScriptLock();

  try {
    writeLog_('DEBUG', 'processPendingBatch_', 'スクリプトロック取得試行中');
    lock.waitLock(30000);
    writeLog_('DEBUG', 'processPendingBatch_', 'スクリプトロック取得成功');

    const startTime = new Date();
    let stepResults = {};

    try {
      // 0. 過去日付のデータを除外
      logBatchProcess_('ステップ0開始', { step: '過去日付データの除外' });
      const archiveResult = archivePastDatePending_();
      stepResults.archive = archiveResult;
      logBatchProcess_('ステップ0完了', { step: '過去日付データの除外', result: archiveResult });

      // 1. 過剰登録のクリーンアップ
      logBatchProcess_('ステップ1開始', { step: '過剰登録のクリーンアップ' });
      const cleanupResult = cleanupOverflowedPending_();
      stepResults.cleanup = cleanupResult;
      logBatchProcess_('ステップ1完了', { step: '過剰登録のクリーンアップ', result: cleanupResult });

      // 2. 空き枠への追加登録
      logBatchProcess_('ステップ2開始', { step: '空き枠への追加登録' });
      const fillResult = fillRemainingSlots_();
      stepResults.fill = fillResult;
      logBatchProcess_('ステップ2完了', { step: '空き枠への追加登録', result: fillResult });

      // 3. pendingのスロットごとの処理
      logBatchProcess_('ステップ3開始', { step: 'pending処理' });
      const pendingResult = processAllPendingSlots_();
      stepResults.pending = pendingResult;
      logBatchProcess_('ステップ3完了', { step: 'pending処理', result: pendingResult });

      // 4. 確定後のデータ整理
      logBatchProcess_('ステップ4開始', { step: '確定後データ整理' });
      const afterResult = cleanupAfterConfirmation();
      stepResults.after = afterResult;
      logBatchProcess_('ステップ4完了', { step: '確定後データ整理', result: afterResult });

      // 5. メール送信
      logBatchProcess_('ステップ5開始', { step: 'メール送信' });
      processMailQueue_();
      logBatchProcess_('ステップ5完了', { step: 'メール送信' });

      const endTime = new Date();
      const duration = endTime.getTime() - startTime.getTime();

      logBatchProcess_('バッチ処理完了', {
        startTime: startTime,
        endTime: endTime,
        durationMs: duration,
        results: stepResults
      });

    } catch (error) {
      logError_('processPendingBatch_', error, {
        stepResults: stepResults,
        timestamp: new Date()
      });
      throw error;
    }

  } finally {
    lock.releaseLock();
    writeLog_('DEBUG', 'processPendingBatch_', 'スクリプトロック解放完了');
  }
}

/** ========= 過去日付のpending/waitlistをArchive ========= */
function archivePastDatePending_() {
  writeLog_('DEBUG', 'archivePastDatePending_', '過去日付データのアーカイブ処理開始');

  // 明日の日付を取得（明日より前 = 今日以前をアーカイブ）
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const tomorrowStr = normDateStr_(tomorrow);

  writeLog_('DEBUG', 'archivePastDatePending_', 'アーカイブ対象日付計算完了', {
    tomorrowStr: tomorrowStr,
    currentDate: new Date()
  });

  const respSh = getSS_().getSheetByName(SHEETS.RESP);
  const responses = respSh.getDataRange().getValues();

  if (responses.length <= 1) {
    writeLog_('INFO', 'archivePastDatePending_', 'アーカイブ対象データなし');
    return { archivedCount: 0, message: 'データなし' };
  }

  const head = responses[0];
  let archivedCount = 0;
  let archivedDetails = [];

  // 後ろから処理（インデックスのズレを防ぐ）
  for (let i = responses.length - 1; i > 0; i--) {
    const row = responses[i];
    const obj = asObj_(head, row);

    try {
      // 日付を確実に正規化して比較
      const objDateStr = normDateStr_(obj.Date);

      // 明日より前（今日以前）の日付のpending/waitlistをArchive
      // confirmedは除外（過去の確定データは別処理で管理）
      if (objDateStr < tomorrowStr && (obj.Status === 'pending' || obj.Status === 'waitlist')) {
        moveToArchive_(obj, 'past-date-pending');
        respSh.deleteRow(i + 1);
        archivedCount++;

        const detail = {
          name: obj.Name,
          email: obj.Email,
          date: objDateStr,
          slotId: obj.SlotID,
          status: obj.Status
        };
        archivedDetails.push(detail);

        writeLog_('DEBUG', 'archivePastDatePending_', '過去日付データをアーカイブ', detail);
      }
    } catch (error) {
      logError_('archivePastDatePending_', error, {
        rowIndex: i,
        obj: obj
      });
    }
  }

  const result = {
    archivedCount: archivedCount,
    details: archivedDetails,
    cutoffDate: tomorrowStr
  };

  if (archivedCount > 0) {
    writeLog_('INFO', 'archivePastDatePending_', `過去日付データアーカイブ完了`, result);
  } else {
    writeLog_('INFO', 'archivePastDatePending_', '過去日付データアーカイブ対象なし');
  }

  return result;
}

function cleanupOverflowedPending_() {
  const conf = readSheetAsObjects_(ensureConfirmedSheet_());
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS));
  
  slots.forEach(slot => {
    const slotId = slot.SlotID;
    const capacity = Number(slot.Capacity);
    const confirmed = conf.find(c => c.SlotID === slotId);
    const actualCount = confirmed ? Number(confirmed.ActualCount || 0) : 0;
    
    if (actualCount >= capacity) {
      const responses = getResponses_().filter(r => 
        r.SlotID === slotId && 
        (r.Status === 'pending' || r.Status === 'waitlist')
      );
      
      responses.forEach(r => {
        moveToArchive_(r, 'slot-already-full');
        deleteResponseRow_(r);
      });
    }
  });
}

function fillRemainingSlots_() {
  const conf = readSheetAsObjects_(ensureConfirmedSheet_());
  const slots = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS));
  
  conf.forEach(confirmed => {
    const actualCount = Number(confirmed.ActualCount || 0);
    const slot = slots.find(s => s.SlotID === confirmed.SlotID);
    if (!slot) return;
    
    const capacity = Number(slot.Capacity);
    const availableSeats = capacity - actualCount;
    
    if (availableSeats > 0 && actualCount >= CONFIG.minCapacityToConfirm) {
      const candidates = getResponses_()
        .filter(r => r.SlotID === confirmed.SlotID && 
                    (r.Status === 'pending' || r.Status === 'waitlist'))
        .sort((a,b) => new Date(a.Timestamp) - new Date(b.Timestamp));
      
      const toAdd = candidates.slice(0, availableSeats).filter(c => 
        !hasConfirmedElsewhere_(c.Email, confirmed.SlotID)
      );
      
      toAdd.forEach(c => {
        setResponseStatus_(c, 'confirmed');
        sendConfirmMail_(c.Name, c.Email, c.Date, c.Start, c.End, slot.Location, slot.Timezone);
      });
      
      if (toAdd.length > 0) {
        updateConfirmedSheet_(confirmed.SlotID);
      }
    }
  });
}

function processAllPendingSlots_() {
  // 明日以降のスロットのみ処理対象とする
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const tomorrowStr = normDateStr_(tomorrow);
  
  const pendingResponses = getResponses_()
    .filter(r => {
      const dateStr = normDateStr_(r.Date);
      return r.Status === 'pending' && dateStr >= tomorrowStr;
    });
    
  const bySlot = groupBy_(pendingResponses, r => r.SlotID);
  
  Object.keys(bySlot).forEach(slotId => {
    confirmIfCapacityReached_(slotId);
  });
}

/** ========= 確定処理 ========= */
function confirmIfCapacityReached_(slotId) {
  const ss = getSS_();
  const slotSh = ss.getSheetByName(SHEETS.SLOTS);
  const respSh = ss.getSheetByName(SHEETS.RESP);

  const slots = readSheetAsObjects_(slotSh);
  const slot = slots.find(s => s.SlotID === slotId);
  if (!slot) return { slotId, status: 'notfound' };
  
  // スロットの日付が明日以降かチェック
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0);
  const tomorrowStr = normDateStr_(tomorrow);
  const slotDateStr = normDateStr_(slot.Date);
  
  if (slotDateStr < tomorrowStr) {
    console.log(`スロット ${slotId} は過去日付のためスキップ: ${slotDateStr}`);
    return { slotId, status: 'past-date', filled: false, confirmedCount: 0 };
  }

  const cap = parseInt(slot.Capacity, 10);
  const minCap = CONFIG.minCapacityToConfirm;

  let all = getResponses_().filter(r => r.SlotID === slotId);
  all.sort((a, b) => new Date(a.Timestamp).getTime() - new Date(b.Timestamp).getTime());

  const allConfirmed = getResponses_().filter(r => r.Status === 'confirmed');
  const confirmedEmails = new Set(allConfirmed.map(r => String(r.Email).toLowerCase()));

  const seenEmailsInThisSlot = new Set();
  const candidates = [];
  
  for (const r of all) {
    if (candidates.length >= cap) break;
    const email = String(r.Email).toLowerCase();
    if (seenEmailsInThisSlot.has(email)) continue;
    if (!CONFIG.allowMultipleConfirmationPerEmail && confirmedEmails.has(email)) continue;
    candidates.push(r);
    seenEmailsInThisSlot.add(email);
  }

  const canConfirm = candidates.length >= minCap;

  const respValues = respSh.getDataRange().getValues();
  const head = respValues.shift();
  const idx = colIndex_(head);

  const newlyConfirmed = [];

  if (!canConfirm) {
    respValues.forEach((row, i) => {
      const obj = asObj_(head, row);
      if (obj.SlotID !== slotId) return;
      if (obj.Status !== 'pending') {
        row[idx.Status] = 'pending';
        row[idx.NotifiedConfirm] = false;
        respSh.getRange(i + 2, 1, 1, row.length).setValues([row]);
      }
    });
    updateSlotAggregate_(slotId, 0, false);
    return { slotId, filled: false, confirmedCount: 0, newlyConfirmed: [] };
  }

  const winners = candidates.slice(0, cap);
  const winnerEmails = new Set(winners.map(w => String(w.Email).toLowerCase()));

  respValues.forEach((row, i) => {
    const obj = asObj_(head, row);
    if (obj.SlotID !== slotId) return;

    const email = String(obj.Email).toLowerCase();
    const isWinner = winnerEmails.has(email);

    if (isWinner) {
      if (obj.Status !== 'confirmed') {
        row[idx.Status] = 'confirmed';
        row[idx.NotifiedWait] = false;
        newlyConfirmed.push({
          rowIndex: i + 2,
          name: obj.Name, 
          email: obj.Email,
          date: obj.Date, 
          start: obj.Start, 
          end: obj.End
        });
      }
    } else {
      if (obj.Status !== 'waitlist') {
        row[idx.Status] = 'waitlist';
      }
    }
    respSh.getRange(i + 2, 1, 1, row.length).setValues([row]);
  });

  const confirmedNowCount = winners.length;

  updateSlotAggregate_(slotId, confirmedNowCount, confirmedNowCount >= cap);

  newlyConfirmed.forEach(nc => {
    sendConfirmMail_(nc.name, nc.email, nc.date, nc.start, nc.end, slot.Location, slot.Timezone);
    markNotified_(nc.rowIndex, 'NotifiedConfirm', true);
  });

  updateConfirmedSheet_(slotId);
  
  if (newlyConfirmed.length > 0) {
    sendAdminConfirmMail_(slot, winners);
  }

  return {
    slotId,
    filled: confirmedNowCount >= cap,
    confirmedCount: confirmedNowCount,
    newlyConfirmed: newlyConfirmed.map(n => n.email)
  };
}

function updateConfirmedSheet_(slotId) {
  const sh = ensureConfirmedSheet_();
  const slot = readSheetAsObjects_(getSS_().getSheetByName(SHEETS.SLOTS))
    .find(s => s.SlotID === slotId);
  if (!slot) return;
  
  const confirmed = getResponses_()
    .filter(r => r.SlotID === slotId && r.Status === 'confirmed')
    .sort((a,b) => new Date(a.Timestamp) - new Date(b.Timestamp));
  
  const rowData = [
    slotId,
    normDateStr_(slot.Date),
    normTimeStr_(slot.Start),
    normTimeStr_(slot.End),
    slot.Location,
    new Date()
  ];
  
  for (let i = 0; i < CONFIG.capacity; i++) {
    if (i < confirmed.length) {
      rowData.push(confirmed[i].Name, confirmed[i].Email);
    } else {
      rowData.push('', '');
    }
  }
  
  rowData.push(confirmed.length);
  
  const values = sh.getDataRange().getValues();
  const head = values.shift();
  const idx = colIndex_(head);
  
  let found = false;
  for (let i = 0; i < values.length; i++) {
    if (values[i][idx.SlotID] === slotId) {
      sh.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
      found = true;
      break;
    }
  }
  
  if (!found) {
    sh.appendRow(rowData);
  }
}

function updateSlotAggregate_(slotId, confirmedCount, filled){
  var sh=getSS_().getSheetByName(SHEETS.SLOTS), vals=sh.getDataRange().getValues(), head=vals.shift(), idx=colIndex_(head), rowIndex=-1;
  for (var i=0;i<vals.length;i++){ 
    if (vals[i][idx.SlotID]===slotId){ 
      rowIndex=i+2; 
      break; 
    } 
  }
  if (rowIndex>0){
    sh.getRange(rowIndex, idx.ConfirmedCount+1).setValue(confirmedCount);
    sh.getRange(rowIndex, idx.Status+1).setValue(filled ? 'filled' : 'open');
  }
}