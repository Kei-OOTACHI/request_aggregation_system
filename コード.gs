// 参考：https://chatgpt.com/share/f89415cd-5567-44c3-93ba-37ed83f2060b

// グローバル変数として定義する定数
const SETTING_SHEET_NAME = "管理者用設定";
const SETTING_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTING_SHEET_NAME);
const SETTING_DATA = SETTING_SHEET.getDataRange().getValues();
const AGGREGATE_SHEET_ID = getFileIdFromUrl(SETTING_DATA[1][0]);  // 集計用スプレッドシートIDに変更
const AGGREGATE_SHEET_NAME = SETTING_DATA[3][0];
const WEBHOOKURL = SETTING_DATA[1][1];
const START_ROW = 8;  // 取得開始行番号を指定
const START_COL = 2;  // 取得開始列番号を指定
const NUM_COLS = 14;  // 取得する列数を指定
const FLAG_COL = 1;  // 集計フラグが存在する列番号
const ID_COL = 16;  // 集計用IDを記入する列番号
const DATE_COL = 15;  // 日付を記入する列番号
const EMAIL_COL = 14;  // メールアドレスを記入する列番号
const TEAM_NAME = 'A2';  // チーム名を取得するためのセル
const SUM_COLS = [7, 8, 9, 10];
const SUM_CELLS = ['P3', 'P4', 'P5', 'P6'];

/**
 * メイン関数。集計フラグに基づいて集計
 */
function aggregateSheetsData() {
  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(300000);  // 5分間ロックを待機

    const userEmail = Session.getActiveUser().getEmail();  // 実行したユーザーのメールアドレスを取得
    const currentDate = new Date();  // 現在の日付を取得

    // スクリプトに紐づいたスプレッドシートを取得
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // シートごとに処理
    const sheets = activeSpreadsheet.getSheets();
    for (const sheet of sheets) {
      // 管理者用設定シートをスキップ
      if (sheet.getName() === SETTING_SHEET_NAME) {
        continue;
      }

      // A2セルの値を取得し、値がない場合は警告を出して集計を中止
      const flagValue = sheet.getRange(TEAM_NAME).getValue();
      if (!flagValue) {
        SpreadsheetApp.getUi().alert(`エラー
        シート "${sheet.getName()}" の${TEAM_NAME}セルに部署名を入力してから再度提出してください。`);
        return;
      }
    }

    // 集計シートの準備
    const aggregateSheet = SpreadsheetApp.openById(AGGREGATE_SHEET_ID).getSheetByName(AGGREGATE_SHEET_NAME);

    // 処理が重くなるのでコメントアウトした
    // 集計シートのクリーニング（すべての値がFALSEまたは空の行を削除）
    // cleanUpAggregateSheet(aggregateSheet);

    const existingIDs = getExistingIDs(aggregateSheet);

    // スプレッドシートを再処理（集計フラグに基づき集計シートに追加・削除）
    processSs(activeSpreadsheet, aggregateSheet, existingIDs, userEmail, currentDate);

    SpreadsheetApp.getUi().alert(`提出完了`);
    // 合計を計算して指定のセルに記入（集計欄が保護されている場合が多く、エラーの温床になっていたのでコメントアウト）
    // for (var i = 0; i < SUM_CELLS.length; i++) {
    //   Logger.log("スタート")
    //   calculateAndWriteSum(activeSpreadsheet, SUM_COLS[i], SUM_CELLS[i]);
    // }
  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert("エラー。申し訳ありません、申請処理中にエラーが発生しました。この画面のスクショとともに大舘にエラーが発生した旨をご一報ください🙇‍♂️\n" + e);
  } finally {
    lock.releaseLock();  // ロックを解除
  }
}

// /**
//  * 集計シートのすべての値がFALSEまたは空の行を削除
//  */
// function cleanUpAggregateSheet(aggregateSheet) {
//   const dataRange = aggregateSheet.getDataRange();
//   const dataValues = dataRange.getValues();

//   // 下から上に向かって行を削除（インデックスの問題を避けるため）
//   for (let i = dataValues.length - 1; i >= 0; i--) {
//     const row = dataValues[i];
//     if (row.every(cell => cell === false || cell === '')) {
//       aggregateSheet.deleteRow(i + 1);
//     }
//   }
// }

/**
 * 各シートのデータを集計フラグに基づいて処理し、集計シートにデータを追加または削除
 */
function processSs(ss, aggregateSheet, existingIDs, userEmail, currentDate) {
  const sheets = ss.getSheets();
  let idCounter = Math.max(...existingIDs) + 1;

  sheets.forEach(sheet => {
    if (sheet.getName() === SETTING_SHEET_NAME) {
      return;
    }

    const flagValue = sheet.getRange(TEAM_NAME).getValue();
    const numRows = sheet.getLastRow() - START_ROW + 1;

    if (numRows < 1) return;

    // シート全体のデータを一括で取得
    const sheetData = sheet.getDataRange().getValues();

    // 必要な部分を slice で抽出
    const flagRangeData = sheetData.slice(START_ROW - 1, START_ROW - 1 + numRows).map(row => [row[FLAG_COL - 1]]);
    const dataRangeData = sheetData.slice(START_ROW - 1, START_ROW - 1 + numRows).map(row => row.slice(START_COL - 1, START_COL - 1 + NUM_COLS));

    const idValues = sheetData.slice(START_ROW - 1, START_ROW - 1 + numRows).map(row => [row[ID_COL - 1]]);
    const dateValues = sheetData.slice(START_ROW - 1, START_ROW - 1 + numRows).map(row => [row[DATE_COL - 1]]);
    const emailValues = sheetData.slice(START_ROW - 1, START_ROW - 1 + numRows).map(row => [row[EMAIL_COL - 1]]);

    // 追加すべきデータをまとめて保持するリスト
    const rowsToAdd = [];

    // 削除すべき集計用IDをまとめて保持するリスト
    const idsToRemove = [];

    for (let i = 0; i < flagRangeData.length; i++) {
      const currentRow = START_ROW + i;
      const existingID = idValues[i][0];  // 集計用IDの有無を確認

      if (flagRangeData[i][0] === true && !existingID) {
        const aggregationID = idCounter.toString();  // 一意の数値のみをIDとして使用

        // メールアドレスと日付の列を新しい値で置換
        dataRangeData[i][EMAIL_COL - START_COL] = userEmail;
        dataRangeData[i][DATE_COL - START_COL] = currentDate;

        // 必要な部分にフラグ値を先頭に追加し、集計用IDを最後に追加
        const rowDataWithFlag = [flagValue, ...dataRangeData[i], aggregationID];

        rowsToAdd.push(rowDataWithFlag);

        // オリジナルシートのメールアドレス、日付、集計用IDを更新
        idValues[i][0] = aggregationID;
        dateValues[i][0] = currentDate;
        emailValues[i][0] = userEmail;

        idCounter++;
      } else if (flagRangeData[i][0] === false && existingID) {
        // 削除対象のIDをリストに追加
        idsToRemove.push(existingID);
        idValues[i][0] = "";
      }
    }

    Logger.log("idValues:" + idValues);
    Logger.log("dateValues:" + dateValues);
    Logger.log("emailValues:" + emailValues);
    // シートへの一括更新
    sheet.getRange(START_ROW, ID_COL, numRows, 1).setValues(idValues);
    sheet.getRange(START_ROW, DATE_COL, numRows, 1).setValues(dateValues);
    sheet.getRange(START_ROW, EMAIL_COL, numRows, 1).setValues(emailValues);

    Logger.log("rowsToAdd" + rowsToAdd);
    // データを一括追加
    addDataToAggregateSheet(rowsToAdd, aggregateSheet);

    Logger.log("idsToRemove" + idsToRemove);
    // 削除対象のIDに基づいて一括削除を実行
    removeDataFromAggregateSheet(idsToRemove, aggregateSheet);

  });
}

/**
 * まとめてデータを集計シートに追加
 */
function addDataToAggregateSheet(rowsToAdd, aggregateSheet) {
  if (rowsToAdd.length === 0) return;  // 追加対象がない場合はスキップ

  const lastRow = aggregateSheet.getLastRow();
  const numRowsToAdd = rowsToAdd.length;

  // 追加する行数分の範囲を取得してデータを一括設定
  aggregateSheet.getRange(lastRow + 1, 1, numRowsToAdd, rowsToAdd[0].length).setValues(rowsToAdd);
}

/**
 * 集計用IDを基にデータを集計シートから削除
 */
function removeDataFromAggregateSheet(idsToRemove, aggregateSheet) {
  if (idsToRemove.length === 0) return;  // 削除対象がない場合はスキップ

  const dataRange = aggregateSheet.getDataRange();
  const dataValues = dataRange.getValues();

  // 削除する行のインデックスを保持するリスト
  const rowsToDelete = [];

  for (let i = 1; i < dataValues.length; i++) {  // ヘッダーをスキップ
    const rowID = dataValues[i][ID_COL - 1];
    if (idsToRemove.includes(rowID)) {
      rowsToDelete.push(i + 1);  // 実際のシート行番号は1始まり
    }
  }

  // 一括削除を実行（行を後ろから削除することでインデックスのずれを防ぐ）
  rowsToDelete.reverse().forEach(rowIndex => {
    aggregateSheet.deleteRow(rowIndex);
  });
}


/**
 * 集計シートから既存の集計用IDを取得
 */
function getExistingIDs(aggregateSheet) {
  // ID_COL列のデータを一括取得
  const idRange = aggregateSheet.getRange(2, ID_COL, aggregateSheet.getLastRow() - 1, 1).getValues();

  // 空白や null を除去して、IDリストを生成
  const existingIDs = idRange
    .flat()  // 2次元配列を1次元に変換
    .filter(id => id);  // 空白やnullの値を削除

  return existingIDs;
}


/**
 * URLからIDを取得する関数
 */
function getFileIdFromUrl(url) {
  let parts = url.split('/');
  let fileId = parts[parts.indexOf("d") + 1];
  return fileId;
}

/**
 * 集計シートの指定列でflagがtrueの行の値を合計し、指定したセルに記入
 */
function calculateAndWriteSum(aggregateSheet, sumCol, sumCell) {
  const dataRange = aggregateSheet.getDataRange();
  const dataValues = dataRange.getValues();
  let sum = 0;

  // 合計を計算
  for (let i = START_ROW - 1; i < dataValues.length; i++) {  // 説明欄と見出し行をスキップ
    Logger.log(dataValues[i])
    if (dataValues[i][FLAG_COL - 1] === true) {
      const value = dataValues[i][sumCol - 1];
      if (typeof value === 'number' && !isNaN(value)) {
        sum += value;
      }
    }
  }

  // 合計を指定のセルに記入
  aggregateSheet.getRange(sumCell).setValue(sum);
}

