// ══════════════════════════════════════════════════════════════
// GAS 新增說明（不要全部取代，只是新增）
//
// 步驟一：在 GAS 編輯器左側，點「+」→「指令碼」
//          把這整個檔案的內容貼入，命名為 GAS_05_assessment
//
// 步驟二：找到原本的 doPost function，在 switch 的 case 區塊裡
//          加入以下這段（放在其他 case 的旁邊，記得有 break）：
//
//   case 'submitAssessment':
//     result = submitAssessmentHandler_(payload);
//     break;
//
// 步驟三：部署新版本
// ══════════════════════════════════════════════════════════════

var ASSESSMENT_SHEET_NAME = '05_經營評量';

function submitAssessmentHandler_(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ASSESSMENT_SHEET_NAME);

    // 自動建立分頁（第一次使用時）
    if (!sheet) {
      sheet = ss.insertSheet(ASSESSMENT_SHEET_NAME);
      // 第1列：英文欄位名
      sheet.getRange(1, 1, 1, 6).setValues([[
        'submitted_at', 'diag_code', 'client_name', 'priority_order', 'content_self_check', 'plan'
      ]]);
      // 第2列：中文欄位名
      sheet.getRange(2, 1, 1, 6).setValues([[
        '送出時間', '查詢代碼', '姓名', '優先順序排序', '內容自評', '合作方案'
      ]]);
      sheet.setFrozenRows(2);
    }

    var diagCode = String(payload.diagCode || '').trim().toUpperCase();
    var priorityOrder = '';
    if (Array.isArray(payload.priorityOrder)) {
      priorityOrder = payload.priorityOrder.map(function(x, i) {
        return (i + 1) + '. ' + x;
      }).join('；');
    }
    var contentSelfCheck = String(payload.contentSelfCheck || '');
    var plan = String(payload.plan || '');

    // 查詢客戶姓名：先找 00_綜合診斷，再找 03_品牌素材庫
    var clientName = '';
    if (diagCode) {
      var diagSheet = ss.getSheetByName('00_綜合診斷');
      if (diagSheet) {
        var dData = diagSheet.getDataRange().getValues();
        for (var i = 2; i < dData.length; i++) {
          if (String(dData[i][1]).trim().toUpperCase() === diagCode) {
            clientName = String(dData[i][2] || '');
            break;
          }
        }
      }
      if (!clientName) {
        var matSheet = ss.getSheetByName('03_品牌素材庫');
        if (matSheet) {
          var mData = matSheet.getDataRange().getValues();
          for (var j = 2; j < mData.length; j++) {
            if (String(mData[j][1]).trim().toUpperCase() === diagCode) {
              clientName = String(mData[j][4] || '');
              break;
            }
          }
        }
      }
    }

    sheet.appendRow([
      new Date(),
      diagCode,
      clientName,
      priorityOrder,
      contentSelfCheck,
      plan
    ]);

    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
