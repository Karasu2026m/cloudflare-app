function myFunction1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_to = ss.getSheetByName('商品マスタ_統合');

  var setdata = [];
  //商品マスタ_統合を消去
  sheet_to.getRange(1, 2, sheet_to.getMaxRows()+10, sheet_to.getMaxColumns()+3).activate();
  ss.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  sheet_copy();
  var data = sheet_to.getRange('B1:AS300').getValues();

  //商品マスタのカテゴリー数が１１個
  for (j = 0;j<11;j++){
    var category = data[0][j*4];
    //各カテゴリ内での行数300あれば足りるかな？
    for (let i = 1;i < 300;i++) {
      const temp = [];
      const str = data[i][j*4];
      if (str.length == 0) {
        break;
      };
      temp.push(data[i][j*4+3],category,data[i][j*4],data[i][j*4+1],data[i][j*4+2],data[i][j*4+3]);
      setdata.push(temp);
    };
  };
  //商品マスタ_統合を消去
  sheet_to.getRange(1, 2, sheet_to.getMaxRows()+10, sheet_to.getMaxColumns()+3).activate();
  ss.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  //データの書き込み
  var endrow = setdata.length;
  sheet_to.getRange('B1:G' + endrow).setValues(setdata);
};

function sheet_copy() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('商品マスタ_参照'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('商品マスタ_統合'), true);
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 2, sheet.getMaxRows(), sheet.getMaxColumns()+1).activate();
  spreadsheet.getRange('\'商品マスタ_参照\'!A1:AR300').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
};

function myFunction() {
  //棚卸しの実施後出庫履歴にコピペするようにデータを抽出する

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_to = ss.getSheetByName('出庫履歴コピペ用');
  var sheet_from = ss.getSheetByName('在庫');
  var setdata = [];

  //var spreadsheet = SpreadsheetApp.getActive();
  sheet_to.getRange('A2:J500').activate();
  sheet_to.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  var data = sheet_from.getRange('A1:Q1000').getValues();
  const str_today = data[0][16];
  for (let i = 1;i < data.length;i++) {
    const temp = [];
    const str = data[i][0];
    if (str.length == 0) {
      break;
    };
    if (data[i][15] == 0){
      continue;
    };
    temp.push(str_today,"",data[i][3],data[i][15],"棚卸し調整");
    setdata.push(temp);
  };

  //前回のデータを消去
  ss.setActiveSheet(ss.getSheetByName('出庫履歴コピペ用'), true);
  sheet_to.getRange("A2:H1000").activate();
  ss.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  //棚卸し調整０件、調整不要だったときの処理、メッセージ

  //データの書き込み
  var endrow = setdata.length+1;
  sheet_to.getRange('A2:E' + endrow).setValues(setdata);

};





