// ============================================================
// ご来店前カウンセリング カルテ保存API（Google Apps Script）
//
// 役割：カウンセリングシート（Webページ）から送られてくる
//       回答＋写真を受け取り、Googleドライブに自動保存する。
//   ・カルテ台帳 … スプレッドシート「カウンセリングカルテ/カルテ台帳」に1送信=1行
//   ・写真       … 「カウンセリングカルテ/写真/日付_名前/」フォルダに保存
//
// デプロイ手順（初回のみ）：
//   1. script.google.com → 新しいプロジェクト → このコードを貼り付け
//   2. プロジェクト名を「カウンセリングカルテAPI」などに変更
//   3. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
//      - 次のユーザーとして実行：自分
//      - アクセスできるユーザー：全員
//   4. 表示された「ウェブアプリのURL」(https://script.google.com/macros/s/…/exec)
//      を index.html の KARTE_ENDPOINT に設定する
// ============================================================

// シート側(index.html)の KARTE_SECRET と一致させる簡易認証キー
const SECRET = 'hashimoto-karte-2026';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.secret !== SECRET) return json_({ ok: false, error: 'unauthorized' });

    // 保存先フォルダ（なければ自動作成）
    const root = getOrCreateFolder_(DriveApp.getRootFolder(), 'カウンセリングカルテ');
    const ss = getOrCreateSpreadsheet_(root, 'カルテ台帳');
    const sheet = ss.getSheets()[0];

    // ヘッダー行（シートが空のときだけ作成）
    if (sheet.getLastRow() === 0) {
      sheet.appendRow((data.headers || []).concat(['写真枚数', '写真フォルダ']));
      sheet.setFrozenRows(1);
    }

    // 写真の保存
    let folderUrl = '';
    const photos = data.photos || [];
    if (photos.length) {
      const photoRoot = getOrCreateFolder_(root, '写真');
      const folder = photoRoot.createFolder((data.dateLabel || '不明日') + '_' + (data.customerName || 'お客様'));
      photos.forEach(function (p, i) {
        const m = String(p.dataUrl || '').match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
        if (!m) return;
        const ext = m[1] === 'image/png' ? '.png' : '.jpg';
        const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], '写真' + (i + 1) + ext);
        folder.createFile(blob);
      });
      folderUrl = folder.getUrl();
    }

    // 台帳に1行追加
    sheet.appendRow((data.row || []).concat([photos.length, folderUrl]));

    return json_({ ok: true, photos: photos.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// 動作確認用（ブラウザでURLを直接開いたとき）
function doGet() {
  return json_({ ok: true, service: 'counseling-karte', time: new Date().toISOString() });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getOrCreateSpreadsheet_(folder, name) {
  const it = folder.getFilesByName(name);
  if (it.hasNext()) return SpreadsheetApp.open(it.next());
  const ss = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder);
  return ss;
}
