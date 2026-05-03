/**
 * summary_dashboard.gs
 * 売上データを月次集計し、サマリーシートへ書き込んだうえで棒グラフを生成する。
 */

// シート名の定数
var SHEET_DATA    = '売上データ';
var SHEET_SUMMARY = '月次サマリー';

// ────────────────────────────────────────────
// メイン処理：集計 → 書き込み → グラフ更新
// ────────────────────────────────────────────
function updateSummaryDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var dataSheet    = ss.getSheetByName(SHEET_DATA);
  var summarySheet = ss.getSheetByName(SHEET_SUMMARY);

  // サマリーシートが存在しない場合は新規作成
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SHEET_SUMMARY);
  }

  // ── 1. 売上データを読み込む ──────────────────
  var lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('売上データが存在しません。');
    return;
  }

  // 2行目以降（ヘッダーを除く）を一括取得
  var values = dataSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // ── 2. 月ごとに集計する ──────────────────────
  // キー：「YYYY年M月」形式の文字列
  var monthMap = {}; // { '2026年1月': { total: 0, count: 0 } }

  values.forEach(function(row) {
    var date   = row[0]; // A列：日付
    var amount = row[3]; // D列：金額

    // 日付が空、または金額が数値でない行はスキップ
    if (!date || typeof amount !== 'number') return;

    var d = new Date(date);
    var label = d.getFullYear() + '年' + (d.getMonth() + 1) + '月';

    if (!monthMap[label]) {
      monthMap[label] = { total: 0, count: 0, year: d.getFullYear(), month: d.getMonth() + 1 };
    }
    monthMap[label].total += amount;
    monthMap[label].count += 1;
  });

  // 年月の昇順に並べ替え
  var sortedLabels = Object.keys(monthMap).sort(function(a, b) {
    var ma = monthMap[a], mb = monthMap[b];
    return ma.year !== mb.year ? ma.year - mb.year : ma.month - mb.month;
  });

  // ── 3. サマリーシートをクリアして書き直す ────
  summarySheet.clearContents();

  // ヘッダー行
  summarySheet.getRange(1, 1, 1, 3).setValues([['月', '合計売上', '件数']]);

  // データ行
  var outputRows = sortedLabels.map(function(label) {
    return [label, monthMap[label].total, monthMap[label].count];
  });
  if (outputRows.length > 0) {
    summarySheet.getRange(2, 1, outputRows.length, 3).setValues(outputRows);
  }

  // 合計売上列を通貨書式に設定
  summarySheet.getRange(2, 2, outputRows.length, 1)
    .setNumberFormat('¥#,##0');

  Logger.log('サマリー書き込み完了：' + outputRows.length + 'ヶ月分');

  // ── 4. 棒グラフを更新する ────────────────────
  updateBarChart(ss, summarySheet, outputRows.length);
}

// ────────────────────────────────────────────
// 棒グラフの作成・更新
// ────────────────────────────────────────────
function updateBarChart(ss, summarySheet, dataRowCount) {
  // 既存のグラフをすべて削除してから再作成
  var existingCharts = summarySheet.getCharts();
  existingCharts.forEach(function(chart) {
    summarySheet.removeChart(chart);
  });

  if (dataRowCount === 0) return;

  // グラフのデータ範囲：A列（月）と B列（合計売上）
  var dataRange = summarySheet.getRange(1, 1, dataRowCount + 1, 2);

  var chart = summarySheet.newChart()
    .setChartType(Charts.ChartType.COLUMN) // 縦棒グラフ
    .addRange(dataRange)
    .setPosition(2, 5, 0, 0)              // 表の右側（E列2行目）に配置
    .setOption('title', '月次売上推移')
    .setOption('hAxis.title', '月')
    .setOption('vAxis.title', '売上金額（円）')
    .setOption('legend.position', 'none')
    .setOption('width', 600)
    .setOption('height', 400)
    .build();

  summarySheet.insertChart(chart);
  Logger.log('グラフの更新が完了しました。');
}

// ────────────────────────────────────────────
// 毎朝9時の自動トリガーを登録する
// ※ GASエディタからこの関数を一度だけ手動実行してください。
//   以降は毎朝9時に updateSummaryDashboard() が自動で起動します。
// ────────────────────────────────────────────
function setDailyTrigger() {
  // 同名のトリガーが重複登録されないよう、既存トリガーを削除する
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'updateSummaryDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎朝 9:00 に実行するトリガーを作成
  ScriptApp.newTrigger('updateSummaryDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('毎朝9時のトリガーを登録しました。');
}
