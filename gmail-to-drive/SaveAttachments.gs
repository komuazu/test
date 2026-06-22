/**
 * Gmail の添付ファイルを Google ドライブの指定フォルダーへ自動保存するスクリプト
 *
 * 動作の流れ:
 *   1. SEARCH_QUERY に一致し、まだ処理していないメールを検索
 *   2. 添付ファイルを DEST_FOLDER_ID のフォルダーへ保存
 *   3. 処理済みラベルを付けて、次回以降の重複保存を防止
 *
 * セットアップ手順は README を参照してください。
 */

// ===== 設定（ここだけ書き換えてください） ============================

// 対象メールの Gmail 検索条件。
//   ・戻ってきた見積PDF:  'in:inbox subject:(見積 OR お見積り) filename:pdf'
//   ・ラベルで絞る場合:    'label:見積回答 filename:pdf'
//   ・差出人で絞る場合:    'in:inbox from:(torihiki@example.com) filename:pdf'
// in:inbox … 自分が送ったものを除き、受信したメールだけを対象にする
// filename:pdf … PDF が付いているメールだけを対象にする
var SEARCH_QUERY = 'in:inbox subject:(見積 OR お見積り) filename:pdf';

// 保存先の既存ドライブフォルダーの ID。
// フォルダーを開いた時の URL  https://drive.google.com/drive/folders/XXXXXXXX の
// XXXXXXXX の部分を貼り付けてください。
var DEST_FOLDER_ID = 'ここにフォルダーIDを貼り付け';

// 処理済みメールに付けるラベル名（重複保存を防ぐための目印。自動で作成されます）
var PROCESSED_LABEL = 'Drive保存済み';

// 同名ファイルが既にフォルダーにある場合に上書きせずスキップするか
var SKIP_IF_EXISTS = true;

// 受信日のサブフォルダー（例 2026-06）を作って整理するか
var ORGANIZE_BY_MONTH = false;

// PDF だけを保存するか（true にすると PDF 以外の添付は無視）
var PDF_ONLY = true;

// ====================================================================

/**
 * メイン処理。トリガーから定期実行されます。
 */
function saveGmailAttachmentsToDrive() {
  var rootFolder = DriveApp.getFolderById(DEST_FOLDER_ID);
  var processedLabel = getOrCreateLabel_(PROCESSED_LABEL);

  // 未処理のものだけを対象にする
  var query = SEARCH_QUERY + ' -label:' + escapeLabelForQuery_(PROCESSED_LABEL);
  var threads = GmailApp.search(query, 0, 50);
  Logger.log('対象スレッド数: ' + threads.length);

  threads.forEach(function (thread) {
    var savedCount = 0;

    thread.getMessages().forEach(function (message) {
      var attachments = message.getAttachments({
        includeInlineImages: false,
        includeAttachments: true
      });

      attachments.forEach(function (attachment) {
        // PDF だけを対象にする設定なら、それ以外はスキップ
        if (PDF_ONLY && !isPdf_(attachment)) {
          return;
        }

        var folder = ORGANIZE_BY_MONTH
          ? getOrCreateSubFolder_(rootFolder, formatMonth_(message.getDate()))
          : rootFolder;

        var name = attachment.getName();
        if (SKIP_IF_EXISTS && folder.getFilesByName(name).hasNext()) {
          Logger.log('スキップ（同名あり）: ' + name);
          return;
        }

        folder.createFile(attachment.copyBlob());
        savedCount++;
        Logger.log('保存: ' + name);
      });
    });

    // スレッド内の添付を処理し終えたら処理済みラベルを付与
    thread.addLabel(processedLabel);
    if (savedCount > 0) {
      Logger.log('スレッド「' + thread.getFirstMessageSubject() + '」: ' + savedCount + ' 件保存');
    }
  });
}

/**
 * 10 分おきの自動実行トリガーを作成します（最初に 1 回だけ実行してください）。
 */
function createTrigger() {
  // 既存の同名トリガーを削除して重複を防止
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'saveGmailAttachmentsToDrive') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('saveGmailAttachmentsToDrive')
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log('トリガーを作成しました（10 分おきに実行）');
}

// ===== 補助関数 =====================================================

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

// 添付が PDF かどうかを MIME タイプまたは拡張子で判定
function isPdf_(attachment) {
  var type = attachment.getContentType() || '';
  var name = (attachment.getName() || '').toLowerCase();
  return type.indexOf('pdf') !== -1 || name.slice(-4) === '.pdf';
}

function getOrCreateSubFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function formatMonth_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
}

// 検索クエリ内でラベル名を使うため、空白を含む場合はクオートする
function escapeLabelForQuery_(name) {
  return /\s/.test(name) ? '"' + name + '"' : name;
}
