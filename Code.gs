/**
 * ============================================================
 * Google Apps Script — Drive Uploader (အပြည့်အစုံ)
 * ============================================================
 * လုပ်ဆောင်ချက်များ:
 *   ၁။ Customer Folder ဖန်တီး
 *   ၂။ ဖိုင်များ တင်
 *   ၃။ အမည်တူ Folder ရှိရင် ငြင်း
 *   ၄။ ၁၂ နာရီပြည့်ရင် Auto Delete
 * ============================================================
 */

// ⚙️ Parent Folder ID (Drive Folder)
const PARENT_FOLDER_ID = "1nEYEbsnysdUNOtI-3EidZwl-nNkCQMIY";

// ⚙️ Access Code
const ACCESS_CODE = "000000";

// ⚙️ Auto Delete ချိန် (နာရီ)
const AUTO_DELETE_HOURS = 12;

// ⚙️ ဖန်တီးချိန် မှတ်ထားတဲ့ Property Key
const TIMESTAMP_KEY = "FOLDER_TIMESTAMPS";

/* ============================================================
   🔵 POST Request — ဖိုင်တင်
   ============================================================ */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ ok: false, message: "ဒေတာ မပါရှိပါ" });
    }

    const data = JSON.parse(e.postData.contents);
    const code = data.code;
    const customerName = (data.customerName || "").trim();
    const files = data.files || [];

    // Access Code စစ်
    if (code !== ACCESS_CODE) {
      return jsonOut({ ok: false, message: "ကုဒ် မှားနေပါသည်" });
    }
    // Customer အမည် စစ်
    if (!customerName) {
      return jsonOut({ ok: false, message: "Customer အမည် မပါရှိပါ" });
    }
    // ဖိုင် စစ်
    if (files.length === 0) {
      return jsonOut({ ok: false, message: "ဖိုင် မပါရှိပါ" });
    }

    // Parent Folder ရှာ
    let parentFolder;
    try {
      parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    } catch (err) {
      return jsonOut({
        ok: false,
        message: "Parent Folder မတွေ့ပါ။ Folder ID ကို စစ်ပါ: " + PARENT_FOLDER_ID
      });
    }

    /* 🔔 အမည်တူ Folder ရှိမရှိ စစ်ဆေး */
    const existing = parentFolder.getFoldersByName(customerName);
    if (existing.hasNext()) {
      const dupFolder = existing.next();
      return jsonOut({
        ok: false,
        duplicate: true,
        message: "အမည်တူနေပါသည်။ အမည် ပြောင်းလဲပါ။",
        existingFolder: dupFolder.getName(),
        existingFolderUrl: dupFolder.getUrl()
      });
    }

    /* Folder အသစ် ဖန်တီး */
    let customerFolder;
    try {
      customerFolder = parentFolder.createFolder(customerName);
    } catch (err) {
      return jsonOut({
        ok: false,
        message: "Folder ဖန်တီးမရပါ: " + String(err.message || err)
      });
    }

    /* ⏱ ဖန်တီးချိန် မှတ် (Auto Delete အတွက်) */
    saveFolderTimestamp(customerFolder.getId(), new Date().getTime());

    /* ဖိုင်များ တင် */
    const uploaded = [];
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      try {
        const bytes = Utilities.base64Decode(f.data);
        const blob = Utilities.newBlob(bytes, f.mimeType, f.name);
        const file = customerFolder.createFile(blob);
        uploaded.push({
          name: file.getName(),
          url: file.getUrl(),
          id: file.getId()
        });
      } catch (fileErr) {
        uploaded.push({
          name: f.name,
          error: String(fileErr.message || fileErr)
        });
      }
    }

    return jsonOut({
      ok: true,
      folder: customerName,
      folderId: customerFolder.getId(),
      folderUrl: customerFolder.getUrl(),
      files: uploaded,
      autoDeleteHours: AUTO_DELETE_HOURS
    });

  } catch (err) {
    return jsonOut({
      ok: false,
      message: "Error: " + String(err.message || err)
    });
  }
}

/* ============================================================
   🟢 GET Request
   ============================================================ */
function doGet(e) {
  return jsonOut({
    ok: true,
    message: "API အလုပ်လုပ်နေပါသည်",
    folderId: PARENT_FOLDER_ID,
    autoDeleteHours: AUTO_DELETE_HOURS
  });
}

/* ============================================================
   ⏱ TIMESTAMP MANAGEMENT
   ============================================================ */
function saveFolderTimestamp(folderId, timeMs) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(TIMESTAMP_KEY);
  let map = {};
  try { map = raw ? JSON.parse(raw) : {}; } catch (e) { map = {}; }
  map[folderId] = timeMs;
  props.setProperty(TIMESTAMP_KEY, JSON.stringify(map));
}

function getFolderTimestamps() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(TIMESTAMP_KEY);
  try { return raw ? JSON.parse(raw) : {}; } catch (e) { return {}; }
}

function setFolderTimestamps(map) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(TIMESTAMP_KEY, JSON.stringify(map));
}

/* ============================================================
   🗑️ AUTO DELETE — ၁၂ နာရီ ကျော်တဲ့ Folder များ ဖျက်
   ============================================================ */
function autoDeleteOldFolders() {
  const startTime = new Date();
  Logger.log("===== Auto Delete စတင်: " + startTime + " =====");

  let parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
  } catch (err) {
    Logger.log("❌ Parent Folder မတွေ့ပါ: " + err.message);
    return;
  }

  const now = new Date().getTime();
  const expiryMs = AUTO_DELETE_HOURS * 60 * 60 * 1000;
  const timestamps = getFolderTimestamps();

  let deletedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  const folders = parentFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    const folderId = folder.getId();
    const folderName = folder.getName();

    let createdAt = timestamps[folderId];
    if (!createdAt) {
      try {
        createdAt = folder.getDateCreated().getTime();
      } catch (e) {
        createdAt = now;
      }
      timestamps[folderId] = createdAt;
    }

    const ageMs = now - createdAt;
    const ageHours = (ageMs / (60 * 60 * 1000)).toFixed(1);

    if (ageMs >= expiryMs) {
      try {
        folder.setTrashed(true);
        delete timestamps[folderId];
        deletedCount++;
        Logger.log("🗑️ ဖျက်ပြီ: " + folderName + " (အသက် " + ageHours + " နာရီ)");
      } catch (err) {
        errorCount++;
        Logger.log("❌ ဖျက်မရ: " + folderName + " — " + err.message);
      }
    } else {
      skippedCount++;
      Logger.log("⏭ ကျန်သေး: " + folderName + " (အသက် " + ageHours + " နာရီ)");
    }
  }

  setFolderTimestamps(timestamps);

  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(2);
  Logger.log("===== ပြီးဆုံး =====");
  Logger.log("🗑️ ဖျက်ပြီး: " + deletedCount);
  Logger.log("⏭ ကျန်သေး: " + skippedCount);
  Logger.log("❌ Error: " + errorCount);
  Logger.log("⏱ အချိန်: " + duration + " စက္ကန့်");

  return {
    deleted: deletedCount,
    skipped: skippedCount,
    errors: errorCount
  };
}

/* ============================================================
   🔧 Helper
   ============================================================ */
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   🧪 Test Functions
   ============================================================ */
function testSetup() {
  try {
    const folder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    Logger.log("✅ Folder တွေ့ပါပြီ: " + folder.getName());
    Logger.log("📁 Folder URL: " + folder.getUrl());
    Logger.log("⏱ Auto Delete: " + AUTO_DELETE_HOURS + " နာရီ");
    return folder.getName();
  } catch (err) {
    Logger.log("❌ Error: " + err.message);
    throw err;
  }
}

function testAutoDelete() {
  const result = autoDeleteOldFolders();
  Logger.log("Result: " + JSON.stringify(result));
}

function testShowTimestamps() {
  const map = getFolderTimestamps();
  Logger.log("Timestamp Map: " + JSON.stringify(map, null, 2));
}

function testClearTimestamps() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(TIMESTAMP_KEY);
  Logger.log("✅ Timestamp အားလုံး ရှင်းပြီ");
}

/* ============================================================
   🚀 TRIGGER SETUP
   ============================================================ */
function createAutoDeleteTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){
    if(t.getHandlerFunction() === 'autoDeleteOldFolders'){
      ScriptApp.deleteTrigger(t);
      Logger.log("🗑️ Trigger အဟောင်း ဖျက်ပြီ");
    }
  });

  ScriptApp.newTrigger('autoDeleteOldFolders')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("✅ Auto Delete Trigger ဖန်တီးပြီ (၁ နာရီတစ်ခါ)");
}

function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){
    Logger.log("Function: " + t.getHandlerFunction() +
               " | Type: " + t.getEventType() +
               " | ID: " + t.getUniqueId());
  });
}

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){ ScriptApp.deleteTrigger(t); });
  Logger.log("🗑️ Trigger အားလုံး ဖျက်ပြီ");
}