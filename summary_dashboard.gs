/**
 * 売上データ 月次サマリー ダッシュボード
 *
 * 「売上データ」シートを月ごとに集計し、「月次サマリー」シートに書き込んだうえで
 * 月次推移の棒グラフを作成する。setDailyTrigger() を一度実行すると、
 * 毎朝 9 時台に runDashboard() が自動実行される。
 *
 * ※ スプレッドシートにバインドされたスクリプト（拡張機能 → Apps Script）として使う想定。
 */

// シート名・設定値
const SOURCE_SHEET_NAME = '売上データ';
const SUMMARY_SHEET_NAME = '月次サマリー';
const SUMMARY_HEADERS = ['月', '合計売上', '件数'];
const TRIGGER_HOUR = 9;
const CHART_TITLE = '月次売上推移';

/**
 * メイン処理：売上データを月ごとに集計し、サマリーとグラフを書き直す。
 * 時間主導型トリガー（毎朝 9 時）から呼び出される。
 */
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    throw new Error('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
  }

  const summary = aggregateByMonth_(sourceSheet, ss.getSpreadsheetTimeZone());
  const summarySheet = getOrCreateSheet_(ss, SUMMARY_SHEET_NAME);

  writeSummary_(summarySheet, summary);
  drawChart_(summarySheet, summary.length);

  console.log('月次サマリーを更新しました（' + summary.length + 'か月分）');
}

/**
 * 毎朝 9 時台に runDashboard を実行するトリガーを登録する。
 * 重複登録を防ぐため、既存の runDashboard トリガーは削除してから作り直す。
 * ※ GAS の仕様上、実行時刻は 9:00〜10:00 の間のいずれかになる。
 * ※ 時刻はプロジェクトのタイムゾーン（appsscript.json の timeZone）に従う。
 */
function setDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function (trigger) {
      return trigger.getHandlerFunction() === 'runDashboard';
    })
    .forEach(function (trigger) {
      ScriptApp.deleteTrigger(trigger);
    });

  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(TRIGGER_HOUR)
    .create();

  console.log('毎朝 ' + TRIGGER_HOUR + ' 時台に runDashboard を実行するトリガーを登録しました');
}

/**
 * 売上データシートの全行を読み込み、月ごとに合計売上と件数を集計する。
 * 1 行目は見出し行として読み飛ばす。日付や金額が不正な行は除外する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 売上データシート
 * @param {string} timeZone 日付を月に変換する際のタイムゾーン
 * @return {Array<Array>} [月ラベル, 合計売上, 件数] の配列（古い月から順）
 */
function aggregateByMonth_(sheet, timeZone) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  // A〜D 列を一括で読み込む（1 行目は見出し）
  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const totals = {}; // キー: 'yyyy-MM'（並べ替え用）、値: { year, month, amount, count }
  let skipped = 0;

  rows.forEach(function (row) {
    const date = toDate_(row[0]);
    const amount = typeof row[3] === 'number' ? row[3] : Number(row[3]);
    if (!date || row[3] === '' || isNaN(amount)) {
      skipped++;
      return;
    }

    const key = Utilities.formatDate(date, timeZone, 'yyyy-MM');
    if (!totals[key]) {
      totals[key] = {
        year: Number(key.slice(0, 4)),
        month: Number(key.slice(5, 7)),
        amount: 0,
        count: 0,
      };
    }
    totals[key].amount += amount;
    totals[key].count++;
  });

  if (skipped > 0) {
    console.warn('日付または金額が不正なため ' + skipped + ' 行をスキップしました');
  }

  return Object.keys(totals)
    .sort()
    .map(function (key) {
      const t = totals[key];
      return [t.year + '年' + t.month + '月', t.amount, t.count];
    });
}

/**
 * セルの値を Date に変換する。日付型でも「2026/01/05」形式の文字列でも受け付ける。
 * @param {*} value セルの値
 * @return {Date|null} 変換できない場合は null
 */
function toDate_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (typeof value === 'string') {
    const m = value.trim().match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) {
      return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    }
  }
  return null;
}

/**
 * 月次サマリーシートをクリアし、見出しと集計結果を書き込む。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 月次サマリーシート
 * @param {Array<Array>} summary 集計結果
 */
function writeSummary_(sheet, summary) {
  // 値・書式と既存のグラフを削除して毎回書き直す
  sheet.clear();
  sheet.getCharts().forEach(function (chart) {
    sheet.removeChart(chart);
  });

  sheet.getRange(1, 1, 1, SUMMARY_HEADERS.length)
    .setValues([SUMMARY_HEADERS])
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  if (summary.length === 0) {
    return;
  }

  // 「2026年1月」が日付に自動変換されないよう、A 列は書式なしテキストにする
  sheet.getRange(2, 1, summary.length, 1).setNumberFormat('@');
  sheet.getRange(2, 1, summary.length, SUMMARY_HEADERS.length).setValues(summary);
  sheet.getRange(2, 2, summary.length, 1).setNumberFormat('#,##0');
  sheet.getRange(2, 3, summary.length, 1).setNumberFormat('0');
  sheet.autoResizeColumns(1, SUMMARY_HEADERS.length);
}

/**
 * 月次サマリーの A〜B 列から月次推移の棒グラフ（縦棒）を作成する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 月次サマリーシート
 * @param {number} dataCount 集計した月数
 */
function drawChart_(sheet, dataCount) {
  if (dataCount === 0) {
    return;
  }

  const chart = sheet.newChart()
    .asColumnChart()
    .addRange(sheet.getRange(1, 1, dataCount + 1, 2)) // 見出し行を含む「月」「合計売上」
    .setNumHeaders(1)
    .setPosition(1, SUMMARY_HEADERS.length + 2, 0, 0) // E1 セルの位置に配置
    .setOption('title', CHART_TITLE)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: '月' })
    .setOption('vAxis', { title: '合計売上（円）', format: '#,##0' })
    .setOption('width', 640)
    .setOption('height', 360)
    .build();

  sheet.insertChart(chart);
}

/**
 * 指定した名前のシートを取得し、存在しなければ新規作成する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss スプレッドシート
 * @param {string} name シート名
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
