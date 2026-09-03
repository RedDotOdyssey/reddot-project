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
const MEMBER_SHEET_NAME = "会员";
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
    if (data.action === "cancelRegistration") {
      return handleCancelRegistration(data);
    }
    if (data.action === "checkin") {
      return handleCheckin(data);
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
      "姓名", "电话", "电邮", "报名人数", "金额(S$)", "签到状态", "活动ID", "支付状态", "报名ID",
    ]);
  } else if (!sheet.getRange(1, 11).getValue()) {
    sheet.getRange(1, 11).setValue("活动ID"); // 兼容旧表格，补上这一列的表头
  }
  if (sheet.getLastRow() > 0 && !sheet.getRange(1, 12).getValue()) {
    sheet.getRange(1, 12).setValue("支付状态"); // 兼容旧表格，补上这一列的表头
  }
  if (sheet.getLastRow() > 0 && !sheet.getRange(1, 13).getValue()) {
    sheet.getRange(1, 13).setValue("报名ID"); // 兼容旧表格，补上这一列的表头
  }

  var eventTitle = String(data.eventTitle || "").trim();
  var eventId = String(data.eventId || "");
  var qty = Math.max(1, Number(data.qty) || 1);
  var cap = Number(data.cap) || 0;

  // 按"活动ID"匹配，而不是按标题文字匹配——避免用"复制模板"创建的、标题相同但
  // 实际是不同场次的活动，被错误地把报名人数加在一起计算，导致误判"名额已满"；
  // 同时要跳过已经标记"已取消"的记录（用统一的 computeRegCounts_ 来算，
  // 保证这里判断"还有没有名额"，跟页面上显示的"已报名人数"永远是同一份依据）
  if (cap > 0 && eventId) {
    var bookedCount = computeRegCounts_()[eventId] || 0;
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
    data.regId || "",
  ]);

  upsertMember_(data.name || "", data.phone || "", data.email || "", eventTitle);

  return jsonOutput({ success: true, message: "报名信息已记录" });
}

/* ---------------- 会员信息（新增，独立于「报名记录」，每人只保留一行） ---------------- */

function getMemberSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MEMBER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MEMBER_SHEET_NAME);
    sheet.appendRow(["姓名", "电话", "电邮", "首次注册时间", "最近一次活动", "最近活动时间", "累计报名次数"]);
  }
  return sheet;
}

// 客人第一次报名时，把姓名/电话/电邮记为一条会员信息；同一个人（按电话或电邮识别，
// 优先用电话——它比电邮更常填、更少打错）之后再报名其他活动，不会重复新增一行，
// 只更新「最近一次活动」和累计报名次数，避免一个人报名多次活动时会员表里出现好几行重复记录
function upsertMember_(name, phone, email, eventTitle) {
  var key = String(phone || "").trim() || String(email || "").trim();
  if (!key) return; // 姓名和电话/电邮都没填全，没法识别是不是同一个人，跳过

  var sheet = getMemberSheet_();
  var rows = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < rows.length; i++) {
    var rowPhone = String(rows[i][1] || "").trim();
    var rowEmail = String(rows[i][2] || "").trim();
    if ((phone && rowPhone === String(phone).trim()) || (email && rowEmail === String(email).trim())) {
      rowIndex = i + 1; // 表格行号从1开始
      break;
    }
  }

  if (rowIndex === -1) {
    // 新会员：第一次报名，记下首次注册时间
    sheet.appendRow([name || "", phone || "", email || "", new Date(), eventTitle || "", new Date(), 1]);
  } else {
    // 老会员：只更新最近一次活动信息和累计次数，首次注册时间不变
    var count = Number(sheet.getRange(rowIndex, 7).getValue()) || 0;
    if (name) sheet.getRange(rowIndex, 1).setValue(name); // 姓名可能后来填得更完整，顺手更新一下
    sheet.getRange(rowIndex, 5).setValue(eventTitle || "");
    sheet.getRange(rowIndex, 6).setValue(new Date());
    sheet.getRange(rowIndex, 7).setValue(count + 1);
  }
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

  // 关键修复：不要直接信任存进 AppData 里的"已报名人数"（那个数字有可能因为
  // 多设备/多次保存时机不同而被旧数据覆盖、跟真实情况脱节）。
  // 每次读取时，都用「报名记录」表里的真实数据重新算一遍，确保显示的数字
  // 永远和真正判断"还有没有名额"时用的是同一份依据，不会再对不上。
  if (payload && payload.events && payload.events.length) {
    var regCounts = computeRegCounts_();
    for (var j = 0; j < payload.events.length; j++) {
      var eid = String(payload.events[j].id);
      if (regCounts.hasOwnProperty(eid)) {
        payload.events[j].reg = regCounts[eid];
      }
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

// 统计「报名记录」表里，每个活动ID对应的真实已报名人数总和
// （这是判断名额够不够时唯一权威的依据，读数据时也用这份，保证两边一致；
//  已标记为"已取消"的行不计入，避免取消了的名额还一直被占着）
function computeRegCounts_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REG_SHEET_NAME);
  var counts = {};
  if (!sheet) return counts;
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][9]) === "已取消") continue;
    var eventId = String(rows[i][10] || "");
    if (!eventId) continue;
    var qty = Number(rows[i][7]) || 0;
    counts[eventId] = (counts[eventId] || 0) + qty;
  }
  return counts;
}

// 客人在 App 里点"取消报名"时调用：找到「报名记录」表里对应的那一行
// （按 活动ID + 姓名 + 电话 匹配，取最近的一条还没被标记取消的记录），
// 把「签到状态」改成"已取消"——不删除原始记录，方便保留历史存档，
// 同时让统计名额时能正确跳过这一笔
function handleCancelRegistration(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REG_SHEET_NAME);
  if (!sheet) return jsonOutput({ success: false, message: "找不到报名记录表" });

  var regId = String(data.regId || "");
  var eventId = String(data.eventId || "");
  var name = String(data.name || "");
  var phone = String(data.phone || "");
  var rows = sheet.getDataRange().getValues();

  // 优先用「报名ID」精确匹配——每一笔报名的编号都是唯一的，不会认错
  if (regId) {
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][12]) === regId && String(rows[i][9]) !== "已取消") {
        sheet.getRange(i + 1, 10).setValue("已取消"); // 第10列 = 签到状态
        return jsonOutput({ success: true, message: "已标记为取消" });
      }
    }
  }

  // 兜底：如果这笔报名是在"报名ID"这个功能上线之前生成的（没有编号可用），
  // 退回用"活动ID + 姓名 + 电话"匹配最近的一条——不如编号精确，但聊胜于无
  for (var j = rows.length - 1; j >= 1; j--) {
    if (
      String(rows[j][10]) === eventId &&
      String(rows[j][4]) === name &&
      String(rows[j][5]) === phone &&
      String(rows[j][9]) !== "已取消"
    ) {
      sheet.getRange(j + 1, 10).setValue("已取消");
      return jsonOutput({ success: true, message: "已标记为取消（按姓名电话匹配，不如报名编号精确，请人工核对一下是否为正确的那一笔）" });
    }
  }
  return jsonOutput({ success: false, message: "未找到匹配的报名记录" });
}

// 扫码签到 / 名单手动签到时调用：按「报名ID」精确匹配「报名记录」表里对应的那一行，
// 把「签到状态」改成"已签到"——这样管理员在 Google Sheet 里看到的状态，
// 才会跟 App 里扫码签到的结果保持一致，不会一直卡在"待签到"
function handleCheckin(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REG_SHEET_NAME);
  if (!sheet) return jsonOutput({ success: false, message: "找不到报名记录表" });

  var regId = String(data.regId || "");
  if (!regId) return jsonOutput({ success: false, message: "缺少报名编号，无法签到" });

  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][12]) === regId) {
      // 已经标记过"已取消"的报名不再改成已签到，避免状态混乱
      if (String(rows[i][9]) !== "已取消") {
        sheet.getRange(i + 1, 10).setValue("已签到"); // 第10列 = 签到状态
      }
      return jsonOutput({ success: true, message: "已标记为已签到" });
    }
  }
  return jsonOutput({ success: false, message: "未找到匹配的报名记录" });
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
