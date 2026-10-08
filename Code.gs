const SHEET_NAME = 'Leaderboard';
const TOP_N = 10;

// 同一個名字只保留最佳一筆；若想改回「每局一列」，把這行改成 false
const KEEP_BEST_PER_NAME = true;

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('找不到名為「' + SHEET_NAME + '」的分頁');
  return sheet;
}

// ===== 接收成績 (POST) =====
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // 等待最多 10 秒，避免多人同時交成績互相覆蓋
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('沒有收到資料');
    }
    const data = JSON.parse(e.postData.contents);

    const name = String(data.playerName || '無名勇者').trim().slice(0, 12);
    const level = Number(data.level) || 0;
    const score = Number(data.score) || 0;

    // 基本防呆：分數異常就不寫入
    if (score < 0 || score > 1000000 || level < 0 || level > 1000) {
      throw new Error('數值超出合理範圍');
    }

    getSheet().appendRow([new Date(), name, level, score]);
    return jsonOut({ status: 'success' });

  } catch (error) {
    return jsonOut({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ===== 讀取排行榜 (GET) =====
function doGet(e) {
  try {
    const rows = getSheet().getDataRange().getValues();
    if (rows.length <= 1) return jsonOut([]);

    // 第 1 列是標題，從第 2 列開始
    let scores = rows.slice(1)
      .filter(row => row[1] !== '' && row[1] !== null)
      .map(row => ({
        name: String(row[1]),
        level: Number(row[2]) || 0,
        score: Number(row[3]) || 0
      }));

    // 先比關卡（高→低），關卡相同再比分數（高→低）
    scores.sort((a, b) => (b.level - a.level) || (b.score - a.score));

    // 同名只留最佳一筆（因為已排序，第一次遇到的就是最好的）
    if (KEEP_BEST_PER_NAME) {
      const seen = {};
      scores = scores.filter(item => {
        if (seen[item.name]) return false;
        seen[item.name] = true;
        return true;
      });
    }

    return jsonOut(scores.slice(0, TOP_N));

  } catch (error) {
    // 回傳空陣列，讓遊戲端不會整個解析失敗
    return jsonOut([]);
  }
}
