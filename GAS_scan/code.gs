/**
 * 💡 v3 完全統合版
 * - v2の機能すべて（入出庫、棚卸、undo、数量・単価変更）を含む
 * - v3の機能（バーコードではなく商品コード記録、不良品報告(上書き＆分割)）を追加
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('📦 在庫スキャンシステム V3')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * 🔒 Web API エンドポイント
 * Cloudflare Workers からのリクエストを受け付ける
 * APIキーで認証し、action名に応じてルーティング
 */
function doPost(e) {
  try {
    var json = JSON.parse(e.postData.contents);
    
    // APIキー検証
    var apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (!apiKey || json.apiKey !== apiKey) {
      return _jsonResponse({ success: false, message: '認証エラー: APIキーが無効です' });
    }
    
    var result;
    switch (json.action) {
      case 'submitBulkCart':
        result = submitBulkCart(json.type, json.cart);
        break;
      case 'getAssemblyOrder':
        result = getAssemblyOrder(json.assemblyId);
        break;
      case 'undoLastScan':
        result = undoLastScan(json.type);
        break;
      case 'ping':
        result = { success: true, message: 'pong', timestamp: getFormattedDate() };
        break;
      default:
        result = { success: false, message: '不明なアクション: ' + json.action };
    }
    
    return _jsonResponse(result);
  } catch (err) {
    return _jsonResponse({ success: false, message: 'サーバーエラー: ' + err.message });
  }
}

/** JSON レスポンスを生成 */
function _jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 現在日時をフォーマット
function getFormattedDate() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function processScan(code, type, defectReason) {
  if(!code) return { success: false, message: "コードが空です" };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // Live sheet is named '在庫' not '在庫_検証'
    var stockSheet = ss.getSheetByName('在庫');
    
    // 👇 棚卸シートがなければ自動で作る
    var targetSheet = ss.getSheetByName(type); 
    if (type === '棚卸') {
      if (!targetSheet) {
        targetSheet = ss.insertSheet('棚卸');
        targetSheet.getRange("A1:G1").setValues([["🎯 スキャン(コード)", "日付", "カテゴリ", "商品名", "システム在庫", "実地の棚卸数", "誤差(過不足)"]]);
        targetSheet.getRange("A1:G1").setBackground("#e1bee7").setFontWeight("bold");
      }
    }
    
    // v3の新機能: 完全に分離されたシートを使用する
    // targetSheetの決定ロジックを上書き
    if (type === '入庫') targetSheet = ss.getSheetByName('入庫_スキャン');
    if (type === '出庫') targetSheet = ss.getSheetByName('出庫_スキャン');
    if (type === '不良報告') targetSheet = ss.getSheetByName('不良在庫');
    // 棚卸はそのまま
    
    // シートが見つからなければエラー（お客様に作ってもらった前提）
    if (!targetSheet) return { success: false, message: "⚠️ 対象シート（" + (type === '入庫' ? '入庫_スキャン' : type === '出庫' ? '出庫_スキャン' : type === '不良報告' ? '不良在庫' : '棚卸') + "）が見つかりません" };

    var masterSheet = ss.getSheetByName('商品マスタ_統合');
    
    // stockSheet（在庫_検証）が無くても、入庫出庫や不良・棚卸は進める方が実用的（一旦許容）
    if (!masterSheet) return { success: false, message: "⚠️ 「商品マスタ_統合」シートが見つかりません" };
    
    // メインシートの商品マスタ_統合の構造
    // A列(0): 在庫管理コード, B列(1): 商品名, C列(2): カテゴリ, D列(3): 商品コード, E列(4): 金額
    var mData = masterSheet.getDataRange().getValues();
    var category = "";
    var productName = "";
    var systemQty = 0;
    var basePrice = 0;
    var found = false;
    
    var searchCode = String(code).trim().toLowerCase().replace(/\s+/g, ' ');
    
    // マスターシートにヘッダーが無い（1行目からデータ）ため j = 0 から検索
    for (var j = 0; j < mData.length; j++) {
      // D列(3):商品コードまたは B列(1):商品名 に対してマッチするかチェック
      var mCode1 = (mData[j].length >= 4 && mData[j][3]) ? String(mData[j][3]).trim().toLowerCase().replace(/\s+/g, ' ') : "";
      var mCode2 = (mData[j].length >= 2 && mData[j][1]) ? String(mData[j][1]).trim().toLowerCase().replace(/\s+/g, ' ') : "";
      var mCode3 = (mData[j].length >= 7 && mData[j][6]) ? String(mData[j][6]).trim().toLowerCase().replace(/\s+/g, ' ') : ""; // G列フォールバック
      
      if (mCode1 === searchCode || mCode2 === searchCode || mCode3 === searchCode) {
        
        productName = mData[j][1] || mData[j][3]; // B列:商品名 または D列:商品コード
        category = mData[j][2] || '未分類';       // C列:カテゴリ
        basePrice = Number(mData[j][4]) || 0;     // E列:単価
        // コードの表記揺れを正すため、マスタ登録の正確な名前に書き換える
        code = mData[j][3] || mData[j][1];
        found = true;
        break;
      }
    }
    
    if(!found) {
        productName = '不明(' + code + ')';
        category = '未分類';
        // return { success: false, message: "⚠️ 未登録の商品です: " + code }; // 強制ブロックしない
    }
    
    // システムの論理在庫取得
    if (stockSheet) {
        // 在庫シートの構造: A:カテゴリ, B:None?, C:在庫管理コード, D:商品コード, E:納期, F:実在庫(5)
        var data = stockSheet.getDataRange().getValues();
        // 在庫シートも1行目（i=0）から探す
        for (var i = 0; i < data.length; i++) {
          if (data[i].length > 3 && (data[i][3] == code || data[i][2] == productName || data[i][2] == code)) {
            systemQty = Number(data[i][5]) || 0; // F列(index 5) = 実在庫
            break;
          }
        }
    }

    var nowText = getFormattedDate();

    // ==========================================
    // ⚠️ 不良報告（振替）モードの処理 (v3)
    // ==========================================
    if (type === '不良報告') {
        let outSheet = ss.getSheetByName('出庫_スキャン');
        if (!outSheet) return { success: false, message: '出庫_スキャンシートが見つかりません' };
        
        let outData = outSheet.getDataRange().getValues();
        let targetRowIndex = -1;
        let targetOriginalQty = 0;
        
        // 下から検索して直近に出庫された同じ商品を特定
        for (let i = outData.length - 1; i >= 1; i--) {
            // 出庫_スキャンシートの新しい構成: [スキャン時間, カテゴリ, 商品名, 商品コード, 数量, 区分, 備考]
            // A列(0):時間, B列(1):カテゴリ, C列(2):商品名, D列(3):商品コード, E列(4):数量, F列(5):区分, G列(6):備考
            let rowItemCode = String(outData[i][3]).trim(); // D列(3)
            let rowItemName = String(outData[i][2]).trim(); // C列(2)
            
            if (rowItemCode === code || rowItemName === productName) {
                let statusStr = '';
                if (outData[i].length > 5) {
                    statusStr = String(outData[i][5]).trim(); // F列(5) = 区分
                }
                
                // 既に不良になっていなければ対象にする
                if(statusStr === '正常' || statusStr === '') {
                     targetRowIndex = i + 1; // 1-indexed
                     targetOriginalQty = Number(outData[i][4]) || 1; // E列(4) = 数量
                     break;
                }
            }
        }
        
        if (targetRowIndex === -1) {
            return { success: false, message: '対象商品の直近の正常出庫履歴が見つかりませんでした。' };
        }
        
        let quantity = 1; // とりあえず1個不良として扱う（その後UIから変更可能）

        // 分割・上書き処理（一旦全数上書きか、1個分だけ切り出すかは、この後の updateLastScanData で数量変更時に本番処理をするのがv2の設計思想だが、一旦ここで「1個不良」として処理する）
        
        // UI操作をスムーズにするために、不良在庫シートに1行追加
        targetSheet.appendRow([nowText, category, productName, 1, defectReason]);
        
        return { success: true, name: productName, qty: 1, price: 0, orderQty: 0, diff: 0, defectRow: targetSheet.getLastRow(), outRow: targetRowIndex, origOutQty: targetOriginalQty };
    }

    // ==========================================
    // 入庫・出庫・棚卸モードの処理 (v2ベース)
    // ==========================================
    var aVals = targetSheet.getRange("A:A").getValues();
    var lastRow = 1;
    for (var k = aVals.length - 1; k >= 0; k--) {
      if (aVals[k][0] !== "") {
        lastRow = k + 1;
        break;
      }
    }
  
    var qtyCol = (type === '入庫') ? 6 : 5; // 入庫_スキャンはF列(6)が入庫数、出庫_スキャンはE列(5)が数量
    var orderQtyCol = 5; // 入庫_スキャンの「発注数」はE列(5)
    
    // ヘッダーが無い場合は自動生成しない（分離シート方式ではユーザーが事前にヘッダーを用意している前提）
    if (type === '出庫') {
      qtyCol = 5;
    }
    
    // --- 棚卸モード専用の処理 ---
    if (type === '棚卸') {
      if (lastRow > 1) {
        var lastCode = targetSheet.getRange(lastRow, 1).getValue();
        if (lastCode == code) {
          var currentAuditQty = Number(targetSheet.getRange(lastRow, 6).getValue()) || 0;
          var newAuditQty = currentAuditQty + 1;
          targetSheet.getRange(lastRow, 6).setValue(newAuditQty); // 実地数を更新
          targetSheet.getRange(lastRow, 7).setValue(newAuditQty - systemQty); // 誤差を更新
          targetSheet.getRange(lastRow, 2).setValue(nowText);
          return { success: true, name: productName, qty: newAuditQty, price: 0, orderQty: 1, diff: newAuditQty - systemQty };
        }
      }
      // 棚卸で新規スキャン
      targetSheet.getRange(lastRow + 1, 1, 1, 7).setValues([[code, nowText, category, productName, systemQty, 1, 1 - systemQty]]);
      return { success: true, name: productName, qty: 1, price: 0, orderQty: 1, diff: 1 - systemQty };
    }
    
  
    // --- 入庫・出庫モードの処理 ---
    if (lastRow > 1) {
      var lastCode = targetSheet.getRange(lastRow, 1).getValue(); 
      if (lastCode == code) {
        // 連続スキャンなら増やす
        var currentQty = Number(targetSheet.getRange(lastRow, qtyCol).getValue()) || 0;
        var newQty = currentQty + 1;
        var currentOrderQty = Number(targetSheet.getRange(lastRow, orderQtyCol).getValue()) || 0;
        targetSheet.getRange(lastRow, qtyCol).setValue(newQty);
        
        if(type === '入庫') targetSheet.getRange(lastRow, orderQtyCol).setValue(currentOrderQty + 1);
        targetSheet.getRange(lastRow, 2).setValue(nowText);
        
        var price = basePrice;
        if (type === '入庫') {
          price = Number(targetSheet.getRange(lastRow, 7).getValue()) || basePrice; 
          targetSheet.getRange(lastRow, 8).setValue(newQty * price);
        }
        return { success: true, name: productName, qty: newQty, price: price, orderQty: currentOrderQty + 1 };
      }
    }
    
    // 直前と違う商品なら新規追加
    if (type === '入庫') {
      // 構成: スキャン時間, カテゴリ, 商品名, 商品コード, 発注数量, 入庫数, 仕入単価, 仕入合計金額
      targetSheet.getRange(lastRow + 1, 1, 1, 8).setValues([[nowText, category, productName, code, 1, 1, basePrice, basePrice]]);
    } else {
      // 出庫の場合: スキャン時間, カテゴリ, 商品名, 商品コード, 数量, 区分, 備考
      targetSheet.getRange(lastRow + 1, 1, 1, 7).setValues([[nowText, category, productName, code, 1, '正常', '']]);
    }
    return { success: true, name: productName, qty: 1, price: basePrice, orderQty: 1 };
  
  } catch(e) {
      return { success: false, message: "エラー: " + e.message };
  }
}


function undoLastScan(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName(type);
  
  if (type === '入庫') targetSheet = ss.getSheetByName('入庫_スキャン');
  if (type === '出庫') targetSheet = ss.getSheetByName('出庫_スキャン');
  // v3: 不良報告のundo対応
  if (type === '不良報告') {
      targetSheet = ss.getSheetByName('不良在庫');
  }
  if (!targetSheet) return { success: false, message: "対象シートがありません" };

  var aVals = targetSheet.getRange("A:A").getValues();
  var lastRow = 1;
  for (var k = aVals.length - 1; k >= 0; k--) {
    if (aVals[k][0] !== "") {
      lastRow = k + 1;
      break;
    }
  }
  
  if(lastRow <= 1) return { success: false, message: "取り消せる履歴がありません" };
  
  var code = targetSheet.getRange(lastRow, 3).getValue(); // C列を通常の商品名とする
  if (type === '不良報告') {
      code = targetSheet.getRange(lastRow, 3).getValue(); // 商品コード(商品名)
      // ※注意: 出庫シート側の「分割」を元に戻す処理は複雑化するため、今回は不良シートの行だけ消す（手動で出庫側の数量を戻す運用が安全）
  }

  targetSheet.deleteRow(lastRow);
  return { success: true, message: code + " の直前の履歴を完全に取り消しました" };
}

// 直前のスキャン数量、価格、【発注数】を上書きする機能
function updateLastScanData(type, code, newQty, newPrice, newOrderQty, defectInfo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName(type);
  
  if (type === '入庫') targetSheet = ss.getSheetByName('入庫_スキャン');
  if (type === '出庫') targetSheet = ss.getSheetByName('出庫_スキャン');
  if (type === '不良報告') targetSheet = ss.getSheetByName('不良在庫');
  
  var aVals = targetSheet.getRange("A:A").getValues();
  var lastRow = 1;
  for (var k = aVals.length - 1; k >= 0; k--) {
    if (aVals[k][0] !== "") {
      lastRow = k + 1;
      break;
    }
  }
  
  if(lastRow <= 1) return { success: false, message: "変更できる履歴がありません" };
  
  if (newQty <= 0) {
    targetSheet.deleteRow(lastRow);
    return { success: true, deleted: true, message: "行を削除しました" };
  }
  
  // ▼ 不良報告（振替）時の振替分割処理
  if (type === '不良報告') {
      // 【重要】ここで出庫シートの分割処理を実行する
      targetSheet.getRange(lastRow, 4).setValue(newQty); // 不良在庫の数量更新
      
      if (defectInfo && defectInfo.outRow && defectInfo.origOutQty) {
          let outSheet = ss.getSheetByName('出庫_スキャン');
          if (newQty === defectInfo.origOutQty) {
              // 全数不良
              outSheet.getRange(defectInfo.outRow, 6).setValue('不良'); // 区分(F)
              outSheet.getRange(defectInfo.outRow, 7).setValue('不良振替'); // 備考(G)
          } else if (newQty < defectInfo.origOutQty) {
              // 一部不良：出庫の数を減らし、不良出庫行を増やす
              outSheet.getRange(defectInfo.outRow, 5).setValue(defectInfo.origOutQty - newQty); // E列:数量を減らす
              let outData = outSheet.getRange(defectInfo.outRow, 1, 1, outSheet.getLastColumn()).getValues()[0];
              outData[4] = newQty; // 新しい数量 (E列)
              outData[5] = '不良'; // 区分 (F列)
              outData[6] = '分割振替'; // 備考 (G列)
              outSheet.appendRow(outData);
          }
      }
      return { success: true, deleted: false, message: "不良数量を更新し、出庫から振替しました" };
  }

  // ▼ 通常処理（入庫・出庫・棚卸）
  var qtyCol = (type === '入庫') ? 6 : 5;
  
  if (type === '棚卸') {
     var sysStock = Number(targetSheet.getRange(lastRow, 5).getValue()) || 0;
     targetSheet.getRange(lastRow, 6).setValue(newQty);
     targetSheet.getRange(lastRow, 7).setValue(newQty - sysStock);
     return { success: true, deleted: false, message: "棚卸の数量を更新しました" };
  }

  targetSheet.getRange(lastRow, qtyCol).setValue(newQty);
  
  if (type === '入庫') {
    var finalOrderQty = (newOrderQty === "" || newOrderQty == null) ? newQty : newOrderQty;
    targetSheet.getRange(lastRow, 5).setValue(finalOrderQty);
    targetSheet.getRange(lastRow, 7).setValue(newPrice);
    targetSheet.getRange(lastRow, 8).setValue(newQty * newPrice);
  }
  
  return { success: true, deleted: false, message: "データを上書き保存しました" };
}

// ==========================================
// V4: バルク処理（カートからのまとめて上書き）
// ==========================================
function submitBulkCart(type, cartArray) {
  if(!cartArray || cartArray.length === 0) return { success: false, message: "カートが空です" };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targetSheet = null;
    
    if (type === '入庫') targetSheet = ss.getSheetByName('入庫_スキャン');
    if (type === '出庫') targetSheet = ss.getSheetByName('出庫_スキャン');
    if (type === '組み立て') targetSheet = ss.getSheetByName('出庫_スキャン'); // 組み立ても実質は出庫
    if (type === '棚卸') targetSheet = ss.getSheetByName('棚卸');
    if (type === '不良報告') targetSheet = ss.getSheetByName('不良在庫');
    
    if (!targetSheet) {
        // 棚卸だけは自動生成
        if (type === '棚卸') {
            targetSheet = ss.insertSheet('棚卸');
            targetSheet.getRange("A1:G1").setValues([["🎯 スキャン(コード)", "日付", "カテゴリ", "商品名", "システム在庫", "実地の棚卸数", "誤差(過不足)"]]);
            targetSheet.getRange("A1:G1").setBackground("#e1bee7").setFontWeight("bold");
        } else {
            return { success: false, message: "対象シートが見つかりません" };
        }
    }
    
    var masterSheet = ss.getSheetByName('商品マスタ_統合');
    var mData = masterSheet ? masterSheet.getDataRange().getValues() : [];
    
    // ▼【高速化】マスタデータを検索キーごとにO(1)アクセス用の連想配列(Map)へ変換
    var masterMap = {};
    for (var j = 0; j < mData.length; j++) {
        let pName = mData[j][1] || mData[j][3]; 
        let pCat = mData[j][2] || '未分類';
        let pPrice = Number(mData[j][4]) || 0;
        let pCode = mData[j][3] || mData[j][1];
        let pInfo = { name: pName, category: pCat, price: pPrice, finalCode: pCode };

        let mCode1 = (mData[j].length >= 4 && mData[j][3]) ? String(mData[j][3]).trim().toLowerCase().replace(/\s+/g, ' ') : "";
        let mCode2 = (mData[j].length >= 2 && mData[j][1]) ? String(mData[j][1]).trim().toLowerCase().replace(/\s+/g, ' ') : "";
        let mCode3 = (mData[j].length >= 7 && mData[j][6]) ? String(mData[j][6]).trim().toLowerCase().replace(/\s+/g, ' ') : "";
        
        if (mCode1) masterMap[mCode1] = pInfo;
        if (mCode2) masterMap[mCode2] = pInfo;
        if (mCode3) masterMap[mCode3] = pInfo;
    }
    
    var stockSheet = ss.getSheetByName('在庫');
    var sData = stockSheet ? stockSheet.getDataRange().getValues() : [];
    
    // ▼【高速化】在庫データをO(1)アクセス用の連想配列(Map)へ変換
    var stockMap = {};
    for (var k = 0; k < sData.length; k++) {
        if (sData[k].length > 3) {
            let sQty = Number(sData[k][5]) || 0;
            if (sData[k][3]) stockMap[String(sData[k][3]).trim().toLowerCase().replace(/\s+/g, ' ')] = sQty;
            if (sData[k][2]) stockMap[String(sData[k][2]).trim().toLowerCase().replace(/\s+/g, ' ')] = sQty;
        }
    }
    
    // ▼【高速化】入庫時の過去履歴から発注数(E列)を探すためのMap生成
    var orderQtyMap = {};
    if (type === '入庫' && targetSheet) {
        var lastRowsData = targetSheet.getDataRange().getValues();
        // 下から上へ走査して最新の発注数を取得
        for (var r = lastRowsData.length - 1; r >= 1; r--) {
            let rName = String(lastRowsData[r][2]).trim(); // C列
            let rCode = String(lastRowsData[r][3]).trim(); // D列
            let rOrderQty = Number(lastRowsData[r][4]) || 0; // E列
            if (rOrderQty > 0) {
                if (rName && !orderQtyMap[rName]) orderQtyMap[rName] = rOrderQty;
                if (rCode && !orderQtyMap[rCode]) orderQtyMap[rCode] = rOrderQty;
            }
        }
    }
    
    var nowText = getFormattedDate();
    var newRows = []; // 一括追加用配列
    
    for (var i = 0; i < cartArray.length; i++) {
        let code = cartArray[i].code;
        let qty = cartArray[i].qty;
        let searchCode = String(code).trim().toLowerCase().replace(/\s+/g, ' ');
        
        // マスター照合 (O(1)検索)
        let productName = cartArray[i].name || '不明(' + code + ')';
        let category = '未分類';
        let basePrice = 0;
        let finalCode = code;
        
        if (masterMap[searchCode]) {
            productName = masterMap[searchCode].name;
            category = masterMap[searchCode].category;
            basePrice = masterMap[searchCode].price;
            finalCode = masterMap[searchCode].finalCode;
        }
        
        // 在庫照合 (O(1)検索)
        let systemQty = stockMap[String(finalCode).trim().toLowerCase().replace(/\s+/g, ' ')] || 
                        stockMap[String(productName).trim().toLowerCase().replace(/\s+/g, ' ')] || 
                        stockMap[searchCode] || 0;
        
        // シート別の行データ生成
        if (type === '入庫') {
            let userOrderQtyForm = cartArray[i].customOrderQty;
            let finalOrderQty = 0;

            if (userOrderQtyForm && String(userOrderQtyForm).trim() !== '') {
                // 手動でお客様が「発注数」を入力した場合はそれを最優先
                finalOrderQty = Number(userOrderQtyForm);
            } else {
                // V3/V4 仕様: 未入力（不明）時は入庫数を発注数（予定数）として適用するか、過去履歴を引き継ぐ
                let currentOrderQty = orderQtyMap[String(productName).trim()] || orderQtyMap[String(finalCode).trim()] || orderQtyMap[searchCode] || 0;
                
                // 過去に履歴があればそれに今回の数を足す。なければ今回の数(=入庫数)を適用
                finalOrderQty = (currentOrderQty > 0) ? currentOrderQty + qty : qty;
            }
            
            // 時間, カテゴリ, 商品名, 商品コード, 発注数量, 入庫数, 仕入単価, 合計
            newRows.push([nowText, category, productName, finalCode, finalOrderQty, qty, basePrice, basePrice * qty]);
        } else if (type === '出庫' || type === '組み立て') {
            // 時間, カテゴリ, 商品名, 商品コード, 数量, 区分, 備考
            let memo = (type === '組み立て') ? '組み立て出庫' : '';
            newRows.push([nowText, category, productName, finalCode, qty, '正常', memo]);
        } else if (type === '棚卸') {
            // スキャン(コード), 日付, カテゴリ, 商品名, システム在庫, 実地の棚卸数, 誤差(過不足)
            newRows.push([finalCode, nowText, category, productName, systemQty, qty, qty - systemQty]);
        } else if (type === '不良報告') {
             // カートUIでの入力項目を取得
             let reason = cartArray[i].defectReason || '初期不良';
             let memoText = cartArray[i].defectMemo || '';
             let defectDesc = reason + (memoText ? ' - ' + memoText : '');
             
             // V4 不良処理（出庫_スキャンからの振替ロジック）
             let outSheet = ss.getSheetByName('出庫_スキャン');
             if (outSheet) {
                 let outData = outSheet.getDataRange().getValues();
                 for (let r = outData.length - 1; r >= 1; r--) {
                     let rowItemCode = String(outData[r][3]).trim(); 
                     let rowItemName = String(outData[r][2]).trim(); 
                     let statusStr = (outData[r].length > 5) ? String(outData[r][5]).trim() : '';
                     
                     if ((rowItemCode === finalCode || rowItemName === productName) && (statusStr === '正常' || statusStr === '')) {
                         let outRow = r + 1;
                         let origOutQty = Number(outData[r][4]) || 0;
                         
                         if (qty >= origOutQty) {
                             // 全数不良（出庫履歴の区分を上書き）
                             outSheet.getRange(outRow, 6).setValue('不良');
                             outSheet.getRange(outRow, 7).setValue('不良振替: ' + defectDesc);
                         } else {
                             // 一部不良：出庫数を減らし、不良行を増やす
                             outSheet.getRange(outRow, 5).setValue(origOutQty - qty);
                             let outDataRow = outSheet.getRange(outRow, 1, 1, outSheet.getLastColumn()).getValues()[0];
                             outDataRow[4] = qty;
                             outDataRow[5] = '不良';
                             outDataRow[6] = '不良振替: ' + defectDesc;
                             outSheet.appendRow(outDataRow);
                         }
                         break; // 1件該当したら抜ける
                     }
                 }
             }
             
             // 不良在庫シートの構成: スキャン時間, カテゴリ, 商品名, 数量, 不良内容
             newRows.push([nowText, category, productName, qty, defectDesc]);
        }
    }
    
    // 一括書き込み (APIのコール回数を減らすため)
    if (newRows.length > 0) {
        let startRow = targetSheet.getLastRow() + 1;
        let cols = newRows[0].length;
        targetSheet.getRange(startRow, 1, newRows.length, cols).setValues(newRows);
    }
    
    return { success: true };
    
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// V4: 組み立て指示書の読み込み（指定カラム・横並び構成に対応）
// ==========================================
function getAssemblyOrder(assemblyId) {
    try {
        // ▼【お客様の設定箇所】別ファイルのURLの「d/」〜「/edit」の間にあるIDをここに貼り付けてください
        var EXTERN_SHEET_ID = 'ここに別スプレッドシートのIDを貼り付けてください';
        
        // ▼【追加】Config.gsからスプレッドシートIDを優先取得する
        try {
            var cfg = getConfig();
            if (cfg && cfg.printSpreadsheetId) {
                EXTERN_SHEET_ID = cfg.printSpreadsheetId;
            }
        } catch(e) { 
            // Config.gsが無い場合はフォールバック(直書きID)で動作させる 
        }

        // もし初期文字列にIDを「付け足して」しまった場合などを考慮し、
        // 実際のIDっぽいもの（長めの英数字）を探すか、そのまま試す
        var sheetIdMatch = EXTERN_SHEET_ID.match(/([a-zA-Z0-9-_]{44,})/);
        var finalId = sheetIdMatch ? sheetIdMatch[1] : EXTERN_SHEET_ID.trim();

        if (!finalId || finalId.includes('ここ')) {
             return { success: false, message: "⚠️ GAS側の EXTERN_SHEET_ID にスプレッドシートのID(URL内の長い英数字)を設定してください" };
        }

        var externSS = SpreadsheetApp.openById(finalId);
        var directiveSheet = externSS.getSheetByName('シート1'); 
        
        if (!directiveSheet) {
             return { success: false, message: "指定された別ファイル内に「シート1」が見つかりません。" };
        }
        
        var data = directiveSheet.getDataRange().getValues();
        var parts = [];
        var foundRow = -1;
        var memoAR = "";
        
        // 組み立て番号が「#」から始まらない場合は補完（念のため）
        var formattedAssemblyId = String(assemblyId).trim();
        if(!formattedAssemblyId.startsWith("#")) {
            formattedAssemblyId = "#" + formattedAssemblyId;
        }
        
        for (var i = 0; i < data.length; i++) {
            if (String(data[i][0]).trim() === formattedAssemblyId) {
                foundRow = i;
                break;
            }
        }
        
        if (foundRow === -1) {
            return { success: false, message: "指定された組み立て番号(" + formattedAssemblyId + ")が見つかりませんでした。" };
        }
        
        var rowData = data[foundRow];
        var partsMap = {};
        
        // お客様指定の「パーツが入っている列」のインデックス (0-indexed)
        // K(10), M(12), O(14), Q(16), S(18), U(20), W(22), Y(24), AA(26), AC(28), AJ(35), AL(37), AN(39), AP(41)
        var targetColumns = [10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 35, 37, 39, 41];
        
        for (var idx = 0; idx < targetColumns.length; idx++) {
            var colIndex = targetColumns[idx];
            
            // 行の長さがその列まで達していない場合はスキップ
            if (colIndex < rowData.length) {
                let partCode = String(rowData[colIndex]).trim();
                if (partCode && partCode !== '') {
                    if (partsMap[partCode]) {
                        partsMap[partCode]++;
                    } else {
                        partsMap[partCode] = 1;
                    }
                }
            }
        }
        
        // AR列(43) の備考を取得
        if (43 < rowData.length) {
            memoAR = String(rowData[43]).trim();
        }

        var masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('商品マスタ_統合');
        var mData = masterSheet ? masterSheet.getDataRange().getValues() : [];
        
        for (var code in partsMap) {
            let searchCode = String(code).trim().toLowerCase().replace(/\s+/g, ' ');
            let pName = code; 
            
            for (var j = 0; j < mData.length; j++) {
                let mCode1 = (mData[j].length >= 4 && mData[j][3]) ? String(mData[j][3]).trim().toLowerCase().replace(/\s+/g, ' ') : "";
                let mCode2 = (mData[j].length >= 2 && mData[j][1]) ? String(mData[j][1]).trim().toLowerCase().replace(/\s+/g, ' ') : "";
                let mCode3 = (mData[j].length >= 7 && mData[j][6]) ? String(mData[j][6]).trim().toLowerCase().replace(/\s+/g, ' ') : "";
                
                if (mCode1 === searchCode || mCode2 === searchCode || mCode3 === searchCode) {
                    pName = mData[j][1] || mData[j][3]; 
                    break;
                }
            }
            
            // 除外リストチェック (pName または code にこれらの文字が含まれていれば飛ばす)
            var omitKeywords = [
                "Win11", "LAN", "PWM", "HUB", "GPUステー", "180Hz", "240Hz", "360Hz", "ナノダイヤモンドグリス", "Office 2021", "Home", "Pro"
            ];
            var shouldOmit = false;
            for (var k = 0; k < omitKeywords.length; k++) {
                if (String(pName).indexOf(omitKeywords[k]) !== -1 || String(code).indexOf(omitKeywords[k]) !== -1) {
                    shouldOmit = true;
                    break;
                }
            }
            
            if (!shouldOmit) {
                parts.push({
                    code: code,
                    name: pName, 
                    requiredQty: partsMap[code]
                });
            }
        }
        
        if (parts.length === 0) {
             return { success: false, message: "該当の行はありましたが、指定されたパーツ列に何も登録されていませんでした。" };
        }
        
        return { success: true, parts: parts, memo: memoAR };
        
    } catch(e) {
        return { success: false, message: "読み込みエラー: " + e.message };
    }
}

