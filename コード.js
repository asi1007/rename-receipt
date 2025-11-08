// ---- 設定値の取得 ----
function getGeminiApiKey() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty("GEMINI_API_KEY");
}

function getRootFolderIds() {
  var props = PropertiesService.getScriptProperties();
  var folderIdsStr = props.getProperty("ROOT_FOLDER_IDS");
  if (!folderIdsStr) return [];
  return folderIdsStr.split(",").map(id => id.trim());
}

function getModelName() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty("MODEL_NAME") || "gemini-2.5-flash";
}

// ---- 初期化関数（初回実行時に1回だけ実行） ----
function initializeProperties() {
  var props = PropertiesService.getScriptProperties();
  
  // 既に設定されている場合はスキップ
  if (props.getProperty("GEMINI_API_KEY")) {
    Logger.log("プロパティは既に設定されています");
    return;
  }
  
  // 初期値を設定
  props.setProperty("GEMINI_API_KEY", "AIzaSyDdlxlaeTR9jAEQjfkiXihsO42EH-okCLw");
  props.setProperty("ROOT_FOLDER_IDS", "1jTDy5xbUXGHRC7M1hTDkeXWWc_6INE2A,1U0ESaGbQV2ICACnzg_SjNXiYh0J-shXP,1S4pl0iPrU0N-852fTrP3unVE1kJqocz,1QJwAqrfEcJ_CpvYAHpaYmfqj8iRNQ21Y");
  props.setProperty("MODEL_NAME", "gemini-2.5-flash");
  
  Logger.log("プロパティを初期化しました");
}

function checkAllPdfsAndRename() {
  var rootFolderIds = getRootFolderIds();
  if (rootFolderIds.length === 0) {
    Logger.log("エラー: ROOT_FOLDER_IDSが設定されていません。initializeProperties()を実行してください。");
    return;
  }
  
  for(let folder_id of rootFolderIds){
    var rootFolder = DriveApp.getFolderById(folder_id);
    processFolder(rootFolder);
  }
}

function processFolder(folder) {
  // フォルダ内のPDFを処理
  var files = folder.getFilesByType(MimeType.PDF);
  while (files.hasNext()) {
    var file = files.next();
    if (!isAlreadyProcessed(file)) {
      processPdf(file);
    }
  }

  // サブフォルダを再帰的に処理
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    processFolder(subfolder);
  }
}

function processPdf(file) {
  var blob = file.getBlob();
  var base64Data = Utilities.base64Encode(blob.getBytes());

  var payload = {
    contents: [{
      parts: [
        {text: `経費精算用のPDFから「請求日」「合計金額」「発行者」「取引物」を抽出してください。
- 請求日は必ずYYMMDD形式。発行日は使用しない。請求日がない場合、発送日は使用してよい
- 発行者は記載されていなければ不明と記載
- 合計金額は消費税込み
- 取引物は請求の対象
- 結果は下記フォーマットでJSONのみで返す。結果の最初にJSONとかつけない。
{"請求日":"YYMMDD","合計金額": xxxx, "発行者":"xxx","取引物":"yyy"}`},
  {inlineData: { mimeType: "application/pdf", data: base64Data }}
      ]
    }]
  };
  
  var modelName = getModelName();
  var apiKey = getGeminiApiKey();
  if (!apiKey) {
    Logger.log("エラー: GEMINI_API_KEYが設定されていません。initializeProperties()を実行してください。");
    return;
  }
  
  var url = `https://generativelanguage.googleapis.com/v1/models/${modelName}:generateContent?key=${apiKey}`;

  var response = UrlFetchApp.fetch(url, 
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );
  Logger.log("もとのテキスト " + response.getContentText());

  var result = JSON.parse(response.getContentText());
  var text = result?.candidates?.[0]?.content?.parts?.[0]?.text || "";
  var jsonTextMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonTextMatch) throw new Error("JSON部分が見つからない");

  try {
    var json = JSON.parse(jsonTextMatch[0]);
    var newName = json["請求日"] + "-" + json["合計金額"] + "-" +json["発行者"] + "-" + json["取引物"] + ".pdf";
    file.setName(newName);
    markAsProcessed(file);
    Logger.log("Renamed: " + newName);
  } catch(e) {
    Logger.log("JSON parse error: " + text);
  }
}

// ---- ユーティリティ ----
function isAlreadyProcessed(file) {
  var props = PropertiesService.getUserProperties();
  return props.getProperty(file.getId()) === "done";
}
function markAsProcessed(file) {
  var props = PropertiesService.getUserProperties();
  props.setProperty(file.getId(), "done");
}
