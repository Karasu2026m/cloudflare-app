/**
 * Config.gs (入出庫管理専用)
 * =========================================
 * 入出庫管理システム設定モジュール
 * Script Propertiesから設定を取得
 * =========================================
 */

/**
 * Script Propertiesから設定を取得して返す
 * @return {Object} 設定オブジェクト
 */
function getConfig() {
  var props = PropertiesService.getScriptProperties();
  
  var config = {
    // --- 入出庫管理設定 ---
    inventorySpreadsheetId: props.getProperty('INVENTORY_SPREADSHEET_ID') || '',
    assemblySheetId:        props.getProperty('ASSEMBLY_SHEET_ID')        || '',

    // --- 一括印刷設定（外部スプレッドシート連携用） ---
    // 入出庫管理側から指示書を発行する場合に使用
    printSpreadsheetId:     props.getProperty('PRINT_SPREADSHEET_ID')     || '1UbaR5QDsLSRk0SU6CB9LvGSEmJKnYHRvDauVOYI78Wc'
  };

  return config;
}
