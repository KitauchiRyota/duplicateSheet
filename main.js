function duplicateSheetAndRename() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  //   ☆ここを編集
  //   コピーする元となるのシート
  const sourceSheetName = "template";
  // コピー後のシート名が格納されているセル範囲
  const sheetNamesRangeAddress = "ポータル!B2:B68";
  //   ここまで☆
  // コピー後のシートのGA名は，D1のセルを更新するようにハードコーディングしてしまっています．変更したい場合は，このコードの50行目あたりを変更してください．

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    Browser.msgBox('エラー', `指定されたシート「${sourceSheetName}」は見つかりませんでした。`, Browser.Buttons.OK);
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
  const portalSheet = ss.getSheetByName('ポータル');
  const startRow = sheetNamesRange.getRow();
  const nameCol = sheetNamesRange.getColumn();

  // 複製処理
  sheetNames.forEach((row, i) => {
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

      // C1セルに新しいシート名を上書き
      newSheet.getRange('D1').setValue(newSheetName);

      // 新しいシートへのリンクを作成し、ポータルシートの隣のセルに書き込む
      const sheetUrl = ss.getUrl() + `#gid=${newSheet.getSheetId()}`;
      portalSheet.getRange(startRow + i, nameCol + 1).setFormula(`=HYPERLINK("${sheetUrl}","${newSheetName}")`);

      Logger.log(`シート「${sourceSheetName}」を「${newSheetName}」として複製し、C1セルに「${newSheetName}」を書き込み、リンクを追加しました。`);
    } catch (e) {
      Browser.msgBox('エラー', `シート「${newSheetName}」の複製中にエラーが発生しました: ${e.message}`, Browser.Buttons.OK);
    }
  });

  Browser.msgBox('完了', 'シートの複製処理が完了しました。', Browser.Buttons.OK);
}