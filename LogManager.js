/** ========= ログ管理機能 ========= */

/**
 * ログを検索・フィルタリングする
 * @param {Object} options - 検索オプション
 * @param {string} options.level - ログレベル (INFO, WARN, ERROR, DEBUG)
 * @param {string} options.function - 関数名
 * @param {Date} options.startDate - 開始日時
 * @param {Date} options.endDate - 終了日時
 * @param {number} options.limit - 取得件数制限
 * @returns {Array} 検索結果
 */
function searchLogs(options = {}) {
  try {
    logFunctionStart_('searchLogs', options);

    const logSheet = ensureLogSheet_();
    const data = logSheet.getDataRange().getValues();

    if (data.length <= 1) {
      return { logs: [], total: 0, message: 'ログデータがありません' };
    }

    const headers = data[0];
    const rows = data.slice(1);

    let filteredRows = rows;

    // レベルフィルター
    if (options.level) {
      filteredRows = filteredRows.filter(row =>
        row[headers.indexOf('Level')] === options.level
      );
    }

    // 関数名フィルター
    if (options.function) {
      filteredRows = filteredRows.filter(row =>
        row[headers.indexOf('Function')].includes(options.function)
      );
    }

    // 日時フィルター
    if (options.startDate || options.endDate) {
      const timestampIndex = headers.indexOf('Timestamp');
      filteredRows = filteredRows.filter(row => {
        const timestamp = new Date(row[timestampIndex]);
        if (options.startDate && timestamp < options.startDate) return false;
        if (options.endDate && timestamp > options.endDate) return false;
        return true;
      });
    }

    // 件数制限
    if (options.limit && options.limit > 0) {
      filteredRows = filteredRows.slice(0, options.limit);
    }

    // オブジェクト形式に変換
    const logs = filteredRows.map(row => {
      const log = {};
      headers.forEach((header, index) => {
        log[header] = row[index];
      });
      return log;
    });

    const result = {
      logs: logs,
      total: filteredRows.length,
      originalTotal: rows.length
    };

    logFunctionEnd_('searchLogs', { resultCount: logs.length });
    return result;

  } catch (error) {
    logError_('searchLogs', error, options);
    throw error;
  }
}

/**
 * 最新のエラーログを取得
 * @param {number} limit - 取得件数
 * @returns {Array} エラーログ
 */
function getRecentErrors(limit = 10) {
  return searchLogs({
    level: 'ERROR',
    limit: limit
  });
}

/**
 * 特定ユーザーのアクションログを取得
 * @param {string} email - ユーザーのメールアドレス
 * @param {number} limit - 取得件数
 * @returns {Array} アクションログ
 */
function getUserActionLogs(email, limit = 20) {
  try {
    const allLogs = searchLogs({ limit: 1000 });

    const userLogs = allLogs.logs.filter(log => {
      try {
        const data = JSON.parse(log.Data || '{}');
        return data.email === email ||
               (data.context && data.context.email === email);
      } catch (e) {
        return false;
      }
    });

    return {
      logs: userLogs.slice(0, limit),
      total: userLogs.length,
      userEmail: email
    };

  } catch (error) {
    logError_('getUserActionLogs', error, { email, limit });
    throw error;
  }
}

/**
 * システムの健康状態をチェック
 * @returns {Object} システム状態
 */
function getSystemHealth() {
  try {
    logFunctionStart_('getSystemHealth');

    const now = new Date();
    const oneHourAgo = new Date(now.getTime() - 60 * 60 * 1000);

    // 過去1時間のログを取得
    const recentLogs = searchLogs({
      startDate: oneHourAgo,
      endDate: now,
      limit: 1000
    });

    const errorCount = recentLogs.logs.filter(log => log.Level === 'ERROR').length;
    const warnCount = recentLogs.logs.filter(log => log.Level === 'WARN').length;
    const batchProcessCount = recentLogs.logs.filter(log =>
      log.Function === 'BATCH_PROCESS' && log.Message.includes('完了')
    ).length;

    // ユーザーアクション数
    const userActionCount = recentLogs.logs.filter(log =>
      log.Function === 'USER_ACTION'
    ).length;

    const health = {
      timestamp: now,
      period: '過去1時間',
      totalLogs: recentLogs.total,
      errorCount: errorCount,
      warningCount: warnCount,
      userActions: userActionCount,
      batchProcesses: batchProcessCount,
      status: errorCount === 0 ? 'HEALTHY' : errorCount < 5 ? 'WARNING' : 'ERROR'
    };

    logFunctionEnd_('getSystemHealth', health);
    return health;

  } catch (error) {
    logError_('getSystemHealth', error);
    return {
      status: 'ERROR',
      message: 'システム状態の取得に失敗しました',
      error: error.message
    };
  }
}

/**
 * ログデータをクリーンアップ（古いログを削除）
 * @param {number} daysToKeep - 保持日数
 * @returns {Object} クリーンアップ結果
 */
function cleanupOldLogs(daysToKeep = 30) {
  try {
    logFunctionStart_('cleanupOldLogs', { daysToKeep });

    const logSheet = ensureLogSheet_();
    const data = logSheet.getDataRange().getValues();

    if (data.length <= 1) {
      return { deletedCount: 0, message: 'クリーンアップ対象なし' };
    }

    const headers = data[0];
    const timestampIndex = headers.indexOf('Timestamp');
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

    let deletedCount = 0;

    // 後ろから処理（インデックスのズレを防ぐ）
    for (let i = data.length - 1; i > 0; i--) {
      const row = data[i];
      const timestamp = new Date(row[timestampIndex]);

      if (timestamp < cutoffDate) {
        logSheet.deleteRow(i + 1);
        deletedCount++;
      }
    }

    const result = {
      deletedCount: deletedCount,
      cutoffDate: cutoffDate,
      daysKept: daysToKeep
    };

    writeLog_('INFO', 'cleanupOldLogs', 'ログクリーンアップ完了', result);
    logFunctionEnd_('cleanupOldLogs', result);

    return result;

  } catch (error) {
    logError_('cleanupOldLogs', error, { daysToKeep });
    throw error;
  }
}

/**
 * ログの統計情報を取得
 * @returns {Object} 統計情報
 */
function getLogStatistics() {
  try {
    logFunctionStart_('getLogStatistics');

    const logSheet = ensureLogSheet_();
    const data = logSheet.getDataRange().getValues();

    if (data.length <= 1) {
      return { total: 0, message: 'ログデータがありません' };
    }

    const headers = data[0];
    const levelIndex = headers.indexOf('Level');
    const functionIndex = headers.indexOf('Function');
    const userIndex = headers.indexOf('User');

    const stats = {
      total: data.length - 1,
      byLevel: {},
      byFunction: {},
      byUser: {},
      oldestLog: null,
      newestLog: null
    };

    // 統計を計算
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // レベル別統計
      const level = row[levelIndex];
      stats.byLevel[level] = (stats.byLevel[level] || 0) + 1;

      // 関数別統計
      const func = row[functionIndex];
      stats.byFunction[func] = (stats.byFunction[func] || 0) + 1;

      // ユーザー別統計
      const user = row[userIndex];
      if (user) {
        stats.byUser[user] = (stats.byUser[user] || 0) + 1;
      }

      // 最古・最新ログ
      const timestamp = new Date(row[headers.indexOf('Timestamp')]);
      if (!stats.oldestLog || timestamp < stats.oldestLog) {
        stats.oldestLog = timestamp;
      }
      if (!stats.newestLog || timestamp > stats.newestLog) {
        stats.newestLog = timestamp;
      }
    }

    logFunctionEnd_('getLogStatistics', { total: stats.total });
    return stats;

  } catch (error) {
    logError_('getLogStatistics', error);
    throw error;
  }
}

/**
 * エラー分析レポートを生成
 * @param {number} days - 分析対象日数
 * @returns {Object} エラー分析結果
 */
function generateErrorReport(days = 7) {
  try {
    logFunctionStart_('generateErrorReport', { days });

    const startDate = new Date();
    startDate.setDate(startDate.getDate() - days);

    const errorLogs = searchLogs({
      level: 'ERROR',
      startDate: startDate,
      limit: 1000
    });

    const report = {
      period: `過去${days}日間`,
      totalErrors: errorLogs.total,
      errorsByFunction: {},
      errorsByMessage: {},
      errorTrends: []
    };

    // エラーを関数別・メッセージ別に集計
    errorLogs.logs.forEach(log => {
      const func = log.Function;
      const message = log.Message;

      report.errorsByFunction[func] = (report.errorsByFunction[func] || 0) + 1;
      report.errorsByMessage[message] = (report.errorsByMessage[message] || 0) + 1;
    });

    // 上位エラーをソート
    report.topErrorFunctions = Object.entries(report.errorsByFunction)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 10);

    report.topErrorMessages = Object.entries(report.errorsByMessage)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 10);

    logFunctionEnd_('generateErrorReport', { totalErrors: report.totalErrors });
    return report;

  } catch (error) {
    logError_('generateErrorReport', error, { days });
    throw error;
  }
}

/** ========= デバッグ支援機能 ========= */

/**
 * デバッグモードでの詳細ログ出力を有効/無効化
 * @param {boolean} enabled - デバッグモードの有効/無効
 */
function setDebugMode(enabled = true) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('DEBUG_MODE', enabled ? 'true' : 'false');

  writeLog_('INFO', 'setDebugMode', `デバッグモード${enabled ? '有効' : '無効'}に設定`, {
    enabled: enabled,
    setBy: Session.getActiveUser().getEmail()
  });

  return { debugMode: enabled, message: `デバッグモードを${enabled ? '有効' : '無効'}にしました` };
}

/**
 * デバッグモードが有効かどうかをチェック
 * @returns {boolean} デバッグモードの状態
 */
function isDebugMode() {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty('DEBUG_MODE') === 'true';
}

/**
 * 特定のユーザーの最新アクティビティを詳細表示
 * @param {string} email - ユーザーメールアドレス
 * @returns {Object} ユーザーアクティビティ詳細
 */
function debugUserActivity(email) {
  try {
    writeLog_('DEBUG', 'debugUserActivity', `ユーザーアクティビティ調査開始`, { targetEmail: email });

    // ユーザーの全アクション取得
    const userActions = getUserActionLogs(email, 50);

    // 現在の申し込み状況
    const responses = getResponses_().filter(r =>
      String(r.Email).toLowerCase() === email.toLowerCase()
    );

    // 最近のエラー（このユーザー関連）
    const recentErrors = searchLogs({
      level: 'ERROR',
      limit: 100
    }).logs.filter(log => {
      try {
        const data = JSON.parse(log.Data || '{}');
        return data.email === email ||
               (data.context && data.context.email === email) ||
               log.Message.includes(email);
      } catch (e) {
        return false;
      }
    });

    const report = {
      userEmail: email,
      timestamp: new Date(),
      actionCount: userActions.total,
      recentActions: userActions.logs.slice(0, 10),
      currentResponses: responses,
      recentErrors: recentErrors.slice(0, 5),
      debugInfo: {
        totalLoggedActions: userActions.total,
        errorCount: recentErrors.length,
        latestActivity: userActions.logs.length > 0 ? userActions.logs[0].Timestamp : null
      }
    };

    writeLog_('DEBUG', 'debugUserActivity', 'ユーザーアクティビティ調査完了', {
      targetEmail: email,
      actionCount: report.actionCount,
      errorCount: recentErrors.length
    });

    return report;

  } catch (error) {
    logError_('debugUserActivity', error, { targetEmail: email });
    throw error;
  }
}

/**
 * システム状態の詳細診断
 * @returns {Object} システム診断結果
 */
function diagnosticSystemStatus() {
  try {
    writeLog_('INFO', 'diagnosticSystemStatus', 'システム診断開始');

    const diagnosis = {
      timestamp: new Date(),
      sheets: {},
      properties: {},
      triggers: [],
      mailQuota: {},
      recentActivity: {}
    };

    // シート状態確認
    const ss = getSS_();
    const sheetNames = Object.values(SHEETS);
    sheetNames.forEach(sheetName => {
      try {
        const sheet = ss.getSheetByName(sheetName);
        diagnosis.sheets[sheetName] = {
          exists: !!sheet,
          rowCount: sheet ? sheet.getLastRow() : 0,
          columnCount: sheet ? sheet.getLastColumn() : 0
        };
      } catch (error) {
        diagnosis.sheets[sheetName] = {
          exists: false,
          error: error.message
        };
      }
    });

    // プロパティ確認
    const properties = PropertiesService.getScriptProperties().getProperties();
    diagnosis.properties = {
      count: Object.keys(properties).length,
      hasSSID: !!properties.SS_ID,
      hasAdminEmails: !!properties.ADMIN_EMAILS,
      debugMode: properties.DEBUG_MODE === 'true'
    };

    // トリガー確認
    const triggers = ScriptApp.getProjectTriggers();
    diagnosis.triggers = triggers.map(trigger => ({
      handlerFunction: trigger.getHandlerFunction(),
      source: trigger.getTriggerSource().toString(),
      eventType: trigger.getEventType().toString()
    }));

    // メール送信状況
    try {
      diagnosis.mailQuota = {
        remaining: MailApp.getRemainingDailyQuota(),
        timestamp: new Date()
      };
    } catch (error) {
      diagnosis.mailQuota = { error: error.message };
    }

    // 最近のアクティビティ
    const recentHealth = getSystemHealth();
    diagnosis.recentActivity = recentHealth;

    writeLog_('INFO', 'diagnosticSystemStatus', 'システム診断完了', {
      sheetsOK: Object.values(diagnosis.sheets).every(s => s.exists),
      propertiesOK: diagnosis.properties.hasSSID && diagnosis.properties.hasAdminEmails,
      triggersCount: diagnosis.triggers.length
    });

    return diagnosis;

  } catch (error) {
    logError_('diagnosticSystemStatus', error);
    return {
      timestamp: new Date(),
      error: error.message,
      status: 'DIAGNOSTIC_FAILED'
    };
  }
}