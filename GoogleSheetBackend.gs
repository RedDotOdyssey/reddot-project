/**
 * 红点时光探索之旅 - App 后端（Google Apps Script）v2
 * ------------------------------------------------------------
 * 这一版新增两个功能，解决"换设备/刷新后活动和公司资料就消失"的问题：
 *
 * 1. 「共享数据同步」：活动列表、公司介绍文字、Logo/相册图片网址，
 *    统一存进这个表格新增的「AppData」工作表里，所有设备读到的都是同一份。
 * 2. 「图片上传」：管理员上传的 Logo / 活动图片 / 公司相册照片，
 *    会被转存到你的 Google Drive 里，返回一个图片网址存进表格
 *    （图片本身不会直接塞进 Sheet 单元格，因为单元格能放的文字有上限，
 *      图片转文字编码后很容易超过这个限制）。
 *
 * 部署 / 更新步骤：
 * 1. 打开你现有的 Google Sheet 对应的 Apps Script 编辑页面
 * 2. 把这个文件的全部内容，覆盖粘贴替换掉原来的代码
 * 3. 点右上角「部署」→「管理部署」→ 点编辑（铅笔）图标 → 「版本」选「新版本」→ 点「部署」
 *    （用这种方式更新，网址不会变，不用再改 App.jsx 里的配置）
 * 4. 第一次运行到 Drive 相关功能时，可能会再次弹出授权窗口，
 *    要求"访问你的 Google 云端硬盘"，按提示允许即可
 *    （这是因为新增了存图片到 Drive 的功能，需要多一项权限）
 */

const APP_DATA_SHEET_NAME = "AppData";
const REG_SHEET_NAME = "报名记录";
const REVIEW_SHEET_NAME = "评价";
const DRIVE_FOLDER_NAME = "红点时光探索之旅_图片";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === "saveAppData") {
      return saveAppData(data.payload);
    }
    if (data.action === "uploadImage") {
      return uploadImageToDrive(data.base64, data.mimeType, data.filename);
    }
    if (data.action === "getAppData") {
      return getAppData();
    }
    if (data.action === "addReview") {
      return handleReview(data);
    }
    if (data.action === "translate") {
      return handleTranslate(data);
    }

    // 没有 action 字段的请求，按"报名信息"处理（原有逻辑）
    return handleRegistration(data);
  } catch (err) {
    return jsonOutput({ success: false, message: "服务器出错：" + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (e.parameter.action === "getAppData") {
    return getAppData(e.parameter.callback);
  }
  return ContentService.createTextOutput(
    "此接口仅接受 POST 请求，用于红点时光探索之旅报名系统。若看到这行字，说明部署已经成功。"
  );
}

/* ---------------- 报名信息（原有功能，不变） ---------------- */

function handleRegistration(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REG_SHEET_NAME) || ss.insertSheet(REG_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "提交时间", "活动名称", "活动日期", "活动地点",
      "姓名", "电话", "电邮", "报名人数", "金额(S$)", "签到状态", "活动ID", "支付状态",
    ]);
  } else if (!sheet.getRange(1, 11).getValue()) {
    sheet.getRange(1, 11).setValue("活动ID"); // 兼容旧表格，补上这一列的表头
  }
  if (sheet.getLastRow() > 0 && !sheet.getRange(1, 12).getValue()) {
    sheet.getRange(1, 12).setValue("支付状态"); // 兼容旧表格，补上这一列的表头
  }

  var eventTitle = String(data.eventTitle || "").trim();
  var eventId = String(data.eventId || "");
  var qty = Math.max(1, Number(data.qty) || 1);
  var cap = Number(data.cap) || 0;

  // 按"活动ID"匹配，而不是按标题文字匹配——避免用"复制模板"创建的、标题相同但
  // 实际是不同场次的活动，被错误地把报名人数加在一起计算，导致误判"名额已满"
  if (cap > 0 && eventId) {
    var rows = sheet.getDataRange().getValues();
    var bookedCount = 0;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][10]) === eventId) {
        bookedCount += Number(rows[i][7]) || 0;
      }
    }
    if (bookedCount + qty > cap) {
      return jsonOutput({
        success: false,
        message: "抱歉，「" + eventTitle + "」仅剩 " + Math.max(0, cap - bookedCount) + " 个名额，报名人数超出剩余名额。",
      });
    }
  }

  sheet.appendRow([
    new Date(), eventTitle, data.eventDate || "", data.eventLocation || "",
    data.name || "", data.phone || "", data.email || "",
    qty, data.price || 0, "待签到", eventId, data.paymentStatus === "pending" ? "待付款" : "已支付",
  ]);

  return jsonOutput({ success: true, message: "报名信息已记录" });
}

/* ---------------- 活动评价（新增，单独存一张可读的表） ---------------- */

function handleReview(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REVIEW_SHEET_NAME) || ss.insertSheet(REVIEW_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["提交时间", "活动名称", "评分", "评价内容", "照片链接"]);
  }

  sheet.appendRow([
    new Date(),
    data.eventTitle || "",
    data.rating || "",
    data.text || "",
    typeof data.photo === "string" ? data.photo : "",
  ]);

  return jsonOutput({ success: true, message: "评价已记录" });
}

/* ---------------- 自动翻译（新增，用 Google Apps Script 自带的翻译服务） ---------------- */

function handleTranslate(data) {
  try {
    var texts = data.texts || {}; // 支持一次传多段文字一起翻译，减少请求次数
    var result = {};
    for (var key in texts) {
      if (!texts[key]) { result[key] = ""; continue; }
      result[key] = LanguageApp.translate(texts[key], "", data.target || "en");
    }
    return jsonOutput({ success: true, translations: result });
  } catch (err) {
    return jsonOutput({ success: false, message: "翻译失败：" + err.toString() });
  }
}

/* ---------------- 共享数据同步（新增） ---------------- */

function getAppDataSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(APP_DATA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(APP_DATA_SHEET_NAME);
    sheet.appendRow(["key", "value", "更新时间"]);
    sheet.appendRow(["shared", "", ""]);
  }
  return sheet;
}

function getAppData(callback) {
  var sheet = getAppDataSheet_();
  var rows = sheet.getDataRange().getValues();
  var payload = null;
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === "shared") {
      var raw = rows[i][1];
      payload = raw ? JSON.parse(raw) : null;
      break;
    }
  }
  var json = JSON.stringify({ success: true, data: payload });

  // JSONP：如果请求里带了 callback 参数（前端用 <script> 标签加载时会带），
  // 就把结果包成一段可执行的 JS 代码返回，绕开"读取跨域内容"被拦截的限制；
  // 没带 callback 的话（比如被当作 POST 那样直接调用），照常返回普通 JSON。
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOutput({ success: true, data: payload });
}

function saveAppData(payload) {
  var sheet = getAppDataSheet_();
  var rows = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === "shared") {
      rowIndex = i + 1; // 表格行号从1开始
      break;
    }
  }
  var json = JSON.stringify(payload);
  if (rowIndex === -1) {
    sheet.appendRow(["shared", json, new Date()]);
  } else {
    sheet.getRange(rowIndex, 2).setValue(json);
    sheet.getRange(rowIndex, 3).setValue(new Date());
  }
  return jsonOutput({ success: true });
}

/* ---------------- 图片上传到 Google Drive（新增） ---------------- */

function getOrCreateFolder_() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function uploadImageToDrive(base64, mimeType, filename) {
  try {
    var folder = getOrCreateFolder_();
    var cleanBase64 = base64.indexOf(",") > -1 ? base64.split(",")[1] : base64; // 去掉 data:image/...;base64, 前缀
    var bytes = Utilities.base64Decode(cleanBase64);
    var blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", filename || "image.jpg");
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 用这种格式的网址，<img> 标签能比较稳定地直接显示 Drive 图片
    var url = "https://lh3.googleusercontent.com/d/" + file.getId();
    return jsonOutput({ success: true, url: url });
  } catch (err) {
    return jsonOutput({ success: false, message: "图片上传失败：" + err.toString() });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
