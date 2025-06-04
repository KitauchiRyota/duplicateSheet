function duplicateSheetAndRename() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheetName = "template"; // 初期値は空文字列
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    Browser.msgBox('エラー', `指定されたシート「${sourceSheetName}」は見つかりませんでした。`, Browser.Buttons.OK);
    return;
  }

  // コピー先のシート名が格納されているセル範囲をユーザーに尋ねる
  const sheetNamesRangeAddress = Browser.inputBox('シート複製', 'コピー先のシート名が格納されているセル範囲を「SheetName!A1:A5」のように入力してください:', Browser.Buttons.OK_CANCEL);
  if (sheetNamesRangeAddress === 'cancel' || sheetNamesRangeAddress === '') {
    Logger.log('ユーザーによりキャンセルされました、またはセル範囲が入力されませんでした。');
    return;
  }

  let sheetNamesRange;
  try {
    sheetNamesRange = ss.getRange(sheetNamesRangeAddress);
  } catch (e) {
    Browser.msgBox('エラー', `入力されたセル範囲「${sheetNamesRangeAddress}」が不正です。エラー: ${e.message}`, Browser.Buttons.OK);
    return;
  }

  const sheetNames = sheetNamesRange.getValues();

  // 複製処理
  sheetNames.forEach(row => {
    const newSheetName = String(row[0]).trim(); // セル値は配列で返されるため、row[0]で値を取得し、Stringに変換してtrim()で空白を除去

    if (newSheetName === '') {
      Logger.log('シート名が空白のためスキップします。');
      return;
    }

    // すでに同じ名前のシートが存在するかチェック
    if (ss.getSheetByName(newSheetName)) {
      Logger.log(`シート「${newSheetName}」は既に存在するためスキップします。`);
      Browser.msgBox('警告', `シート「${newSheetName}」は既に存在するため、このシートの複製はスキップされました。`, Browser.Buttons.OK);
      return;
    }

    try {
      const newSheet = sourceSheet.copyTo(ss);
      newSheet.setName(newSheetName);

      // A3セルに新しいシート名を上書き
      newSheet.getRange('A3').setValue(newSheetName);

      Logger.log(`シート「${sourceSheetName}」を「${newSheetName}」として複製し、A3セルに「${newSheetName}」を書き込みました。`);
    } catch (e) {
      Browser.msgBox('エラー', `シート「${newSheetName}」の複製中にエラーが発生しました: ${e.message}`, Browser.Buttons.OK);
    }
  });

  Browser.msgBox('完了', 'シートの複製処理が完了しました。', Browser.Buttons.OK);
}