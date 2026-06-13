const SPREADSHEET_ID = "12tiBlIUli3cYGdLCcE8XS3t8VJsYtyB9qrgLBddMA88";
const SHEET_NAME = "就職活動";

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      const sheets = ss.getSheets().map(s => s.getName());
      return ContentService.createTextOutput(JSON.stringify({ error: "sheet not found", available: sheets })).setMimeType(ContentService.MimeType.JSON);
    }
    const lastRow = sheet.getLastRow();
    const values = sheet.getRange(1, 1, lastRow, 2).getValues();
    let jobTypeTableStart = lastRow + 1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === "Row ID" && values[i][1] === "職種名") {
        jobTypeTableStart = i + 1;
        break;
      }
    }
    let lastDataRow = 1;
    for (let i = 0; i < jobTypeTableStart - 1; i++) {
      if (values[i][0] !== "") lastDataRow = i + 1;
    }
    return ContentService.createTextOutput(JSON.stringify({
      sheetName: sheet.getName(),
      lastRow: lastRow,
      jobTypeTableStart: jobTypeTableStart,
      lastDataRow: lastDataRow,
      targetRow: lastDataRow + 1
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
    const data = JSON.parse(e.postData.contents);
    const result = appendJobData(sheet, data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", writtenRow: result.row, id: result.id }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Claude が使えない場合の手動追加用
function addJobManually() {
  const jsonString = `ここにClaudeが出力したJSONを貼り付け`;
  const data = JSON.parse(jsonString);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  appendJobData(sheet, data);
}

function appendJobData(sheet, data) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 2).getValues();

  // 職種名テーブルの開始行を探す（Col A = "Row ID" かつ Col B = "職種名"）
  let jobTypeTableStart = lastRow + 1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === "Row ID" && values[i][1] === "職種名") {
      jobTypeTableStart = i + 1;
      break;
    }
  }

  // 職種名テーブルより前の、最後に値が入っている行を探す
  let lastDataRow = 1;
  for (let i = 0; i < jobTypeTableStart - 1; i++) {
    if (values[i][0] !== "") lastDataRow = i + 1;
  }
  const targetRow = lastDataRow + 1;

  // 書き込み先が職種名テーブルと重なる場合は行を挿入してスペースを確保
  if (targetRow >= jobTypeTableStart) {
    sheet.insertRowAfter(lastDataRow);
  }

  // 既存データと同形式の8桁 hex ID を生成
  const newId = Utilities.getUuid().replace(/-/g, "").substring(0, 8);

  sheet.getRange(targetRow, 1, 1, 24).setValues([[
    newId,
    data["企業名"]       || "",
    data["雇用形態"]      || "",
    data["応募媒体"]      || "",
    data["職種"]         || "",
    data["応募日"]       || "",
    data["状態"]         || "",
    data["書類選考"]      || "",
    data["一次選考"]      || "",
    data["一次選考合否"]   || "",
    data["二次選考"]      || "",
    data["二次選考合否"]   || "",
    data["内定"]         || "",
    data["給与"]         || "",
    data["業務内容"]      || "",
    data["勤務地"]       || "",
    data["勤務時間"]      || "",
    data["休日数"]       || "",
    data["求人URL"]      || "",
    data["企業URL"]      || "",
    data["備考"]         || "",
    data["総合評価"]      || "",
    data["従業員数"]      || "",
    data["選考の流れ"]    || ""
  ]]);
  return { row: targetRow, id: newId };
}
