/**
 * 红点时光探索之旅 - 报名数据后端（Google Apps Script）
 * ------------------------------------------------------------
 * 用途：接收 App 发来的报名信息，写入 Google Sheet，并做基本的名额校验，
 *      防止两人同时抢最后一个名额时都报名成功（超卖）。
 *
 * 部署步骤：
 * 1. 打开 https://sheets.google.com ，新建一个空白表格，命名为「红点时光报名记录」
 * 2. 菜单栏「扩展程序」→「Apps Script」，会打开一个新的代码编辑窗口
 * 3. 把这个文件里的全部代码复制粘贴进去，覆盖默认的 myFunction() 内容
 * 4. 点右上角「部署」→「新建部署」
 *    - 类型选择「Web 应用」
 *    - 说明随便填，比如「报名接口 v1」
 *    - 「执行身份」选「我」（你自己的 Google 账号）
 *    - 「谁可以访问」选「任何人」（这样 App 才能不登录就调用）
 * 5. 点「部署」，第一次会要求你授权，按提示允许即可
 * 6. 部署成功后会给你一个网址，形如：
 *    https://script.google.com/macros/s/AKfycb.../exec
 *    这个网址就是要填进 App.jsx 里 GOOGLE_SHEET_WEBHOOK_URL 常量的值
 *
 * 以后如果改了代码，要「新建部署」一次新版本，网址才会生效最新代码
 * （或者用「管理部署」→ 编辑 → 更新现有部署也可以，网址不会变）
 */

function doPost(e) {
  // 用 LockService 给写入操作加锁，确保同一时间只有一个报名在处理，
  // 避免两人几乎同时报名时，都读到"名额还没满"从而一起超卖
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); // 最多等待 15 秒

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("报名记录") || ss.insertSheet("报名记录");

    // 空表先写表头
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "提交时间", "活动名称", "活动日期", "活动地点",
        "姓名", "电话", "电邮", "报名人数", "金额(S$)", "签到状态",
      ]);
    }

    var data = JSON.parse(e.postData.contents);

    var eventTitle = String(data.eventTitle || "").trim();
    var qty = Math.max(1, Number(data.qty) || 1);
    var cap = Number(data.cap) || 0;

    // ---- 名额校验：统计该活动目前已登记的总人数，防止超卖 ----
    if (cap > 0 && eventTitle) {
      var rows = sheet.getDataRange().getValues();
      var bookedCount = 0;
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][1]).trim() === eventTitle) {
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
      new Date(),
      eventTitle,
      data.eventDate || "",
      data.eventLocation || "",
      data.name || "",
      data.phone || "",
      data.email || "",
      qty,
      data.price || 0,
      "待签到",
    ]);

    return jsonOutput({ success: true, message: "报名信息已记录" });
  } catch (err) {
    return jsonOutput({ success: false, message: "服务器出错：" + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// 浏览器直接打开这个网址时显示的提示（方便你确认部署是否成功）
function doGet(e) {
  return ContentService.createTextOutput(
    "此接口仅接受 POST 请求，用于红点时光探索之旅报名系统。若看到这行字，说明部署已经成功。"
  );
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
