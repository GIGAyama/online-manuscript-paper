/**
 * オンライン原稿用紙 - サーバーサイドロジック (Code.gs)
 * * このファイルは、Googleスプレッドシートと連携してデータの保存や読み込みを行います。
 * 先生が管理しやすいよう、設定や処理の内容を日本語で説明しています。
 */

// ==========================================
//  1. 設定と定数 (ここを変更すると全体が変わります)
// ==========================================

// 利用するシートの名前
const SHEET_NAME = '作文データ';

// 列（カラム）の定義
// スプレッドシートの何列目に何のデータが入るかを決めています。
// 列番号は A列=1, B列=2, ... と数えます。
const COLUMNS = {
  ID: 1,          // A列: データのID（システムが使う識別番号）
  TITLE: 2,       // B列: 題名
  CLASS: 3,       // C列: 学年・クラス
  NAME: 4,        // D列: 氏名
  CONTENT: 5,     // E列: 作文の本文
  CREATED_AT: 6,  // F列: 作成日時
  UPDATED_AT: 7,  // G列: 更新日時
  DELETED_AT: 8   // H列: 削除日時（これが入っているとゴミ箱扱い）
};

// ==========================================
//  2. 画面を表示する処理
// ==========================================

/**
 * Webアプリにアクセスした時に最初に呼ばれる関数
 */
function doGet() {
  // 'index.html' ファイルを読み込んでWebページとして表示します
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('オンライン原稿用紙')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0') // スマホ対応
      .setFaviconUrl('https://drive.google.com/uc?id=1EsaLbGPFc9WixYhJ5sPynTIBZpxzSsfK&.png');
}

// ==========================================
//  3. データの保存・更新処理
// ==========================================

/**
 * 作文を保存または更新します。
 * 40人一斉アクセスでもデータが壊れないよう「排他制御（ロック）」を行っています。
 * * @param {Object} draftData - クライアントから送られてきた作文データ
 */
function saveOrUpdateDraft(draftData) {
  // ■ ロック処理を開始（同時書き込み防止）
  const lock = LockService.getScriptLock();
  try {
    // 最大10秒間、他の人の処理が終わるのを待ちます
    lock.waitLock(10000); 
  } catch (e) {
    // 10秒待っても空かなかった場合のエラー
    return { status: 'error', message: 'ただいま混み合っています。もう一度「保存」ボタンを押してください。' };
  }

  try {
    const sheet = getSheet_();
    const now = new Date();

    // IDがある場合は「更新」、ない場合は「新規作成」
    if (draftData.id) {
      // --- 更新処理 ---
      const foundRow = findRowById_(sheet, draftData.id);
      
      if (foundRow > 0) {
        // 更新するデータを準備（作成日時は変更しない）
        // getRange(行, 列) でセルを指定して書き込みます
        sheet.getRange(foundRow, COLUMNS.TITLE).setValue(draftData.title);
        sheet.getRange(foundRow, COLUMNS.CLASS).setValue(draftData.class);
        sheet.getRange(foundRow, COLUMNS.NAME).setValue(draftData.name);
        sheet.getRange(foundRow, COLUMNS.CONTENT).setValue(draftData.content);
        sheet.getRange(foundRow, COLUMNS.UPDATED_AT).setValue(now);
        sheet.getRange(foundRow, COLUMNS.DELETED_AT).setValue(''); // ゴミ箱から戻す場合のためにクリア

        return { status: 'success', message: '上書き保存しました。', id: draftData.id };
      }
      // IDが送られてきたが見つからない場合は、念のため新規作成として扱います
    }
    
    // --- 新規作成処理 ---
    const newId = Utilities.getUuid(); // 新しいIDを発行
    
    // appendRowで一番下の行に追加します
    // 配列の順番は COLUMNS の定義と一致させる必要があります
    sheet.appendRow([
      newId,              // A列
      draftData.title,    // B列
      draftData.class,    // C列
      draftData.name,     // D列
      draftData.content,  // E列
      now,                // F列 (作成日)
      now,                // G列 (更新日)
      ''                  // H列 (削除フラグは空)
    ]);
    
    return { status: 'success', message: '新しく保存しました。', id: newId };

  } catch (e) {
    // 予期せぬエラーが起きた場合
    Logger.log('Save Error: ' + e.toString());
    return { status: 'error', message: '保存に失敗しました。先生に伝えてください。\n(' + e.message + ')' };
  } finally {
    // ■ ロックを解除（必ず行う）
    lock.releaseLock(); 
  }
}

// ==========================================
//  4. データの読み込み・一覧取得
// ==========================================

/**
 * 保存されている作文の一覧を取得します。
 * 削除されたもの（ゴミ箱行き）は除外します。
 * 最新の50件のみを返します（動作を軽くするため）。
 */
function getDraftList() {
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; // データがない場合

    // データ範囲を一括取得（2行目から最終行まで、全カラム）
    // values[行][列] の形でデータが入っています（列は0始まりになることに注意）
    const range = sheet.getRange(2, 1, lastRow - 1, 8);
    const values = range.getValues();

    // 以下の条件でデータを加工します
    // 1. 削除日時(8列目/インデックス7)が入っていないものだけ選ぶ
    // 2. 必要なデータだけ抜き出す
    const validDrafts = values
      .filter(row => row[COLUMNS.DELETED_AT - 1] === '') 
      .map(row => ({
        id: row[COLUMNS.ID - 1],
        title: row[COLUMNS.TITLE - 1],
        name: row[COLUMNS.NAME - 1],
        updatedAtRaw: new Date(row[COLUMNS.UPDATED_AT - 1]) // 並び替え用の日付データ
      }));

    // 更新日時が新しい順に並び替え
    validDrafts.sort((a, b) => b.updatedAtRaw - a.updatedAtRaw);

    // 最新50件に絞り、表示用の日付フォーマットを整える
    return validDrafts.slice(0, 50).map(d => ({
      id: d.id,
      title: d.title,
      name: d.name,
      updatedAt: Utilities.formatDate(d.updatedAtRaw, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
    }));

  } catch (e) {
    Logger.log('GetList Error: ' + e.toString());
    return []; // エラー時は空リストを返す
  }
}

/**
 * 指定されたIDの作文データを詳しく読み込みます。
 */
function loadDraft(id) {
  try {
    const sheet = getSheet_();
    const rowIndex = findRowById_(sheet, id);

    if (rowIndex > 0) {
      // その行のデータを取得
      const rowData = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];

      // 削除済みチェック
      if (rowData[COLUMNS.DELETED_AT - 1] !== '') {
        return { status: 'error', message: 'この作文はゴミ箱に入っています。' };
      }

      return {
        status: 'success',
        data: {
          id: rowData[COLUMNS.ID - 1],
          title: rowData[COLUMNS.TITLE - 1],
          class: rowData[COLUMNS.CLASS - 1],
          name: rowData[COLUMNS.NAME - 1],
          content: rowData[COLUMNS.CONTENT - 1]
        }
      };
    }
    return { status: 'error', message: 'データが見つかりませんでした。' };
  } catch (e) {
    Logger.log('Load Error: ' + e.toString());
    return { status: 'error', message: '読み込みに失敗しました。' };
  }
}

// ==========================================
//  5. 削除（アーカイブ）処理
// ==========================================

/**
 * 作文を論理削除（ゴミ箱へ移動）します。
 * データは消さずに、削除日時に日付を入れて見えなくします。
 */
function deleteDraft(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); // 5秒待機
  } catch (e) {
    return { status: 'error', message: 'ただいま混み合っています。もう一度ボタンを押してください。' };
  }

  try {
    const sheet = getSheet_();
    const rowIndex = findRowById_(sheet, id);

    if (rowIndex > 0) {
      // 削除日時カラム(H列)に現在日時を書き込む
      sheet.getRange(rowIndex, COLUMNS.DELETED_AT).setValue(new Date());
      return { status: 'success', message: '作文をゴミ箱へ移動しました。' };
    }
    return { status: 'error', message: '削除するデータが見つかりませんでした。' };

  } catch (e) {
    Logger.log('Delete Error: ' + e.toString());
    return { status: 'error', message: '削除できませんでした。先生に伝えてください。' };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
//  6. ユーティリティ関数（裏方の処理）
// ==========================================

/**
 * シートオブジェクトを取得する関数
 * シートが見つからない場合のエラー処理を共通化しています
 */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`シート「${SHEET_NAME}」が見つかりません。シート名を確認してください。`);
  }
  return sheet;
}

/**
 * IDから行番号を高速に検索する関数
 * 配列ループではなく、GASの検索機能(TextFinder)を使うため高速です。
 * * @param {Sheet} sheet - 検索対象のシート
 * @param {string} id - 検索するID
 * @return {number} 行番号 (見つからない場合は -1)
 */
function findRowById_(sheet, id) {
  // A列(ID列)の中から検索
  const textFinder = sheet.getRange("A:A").createTextFinder(id);
  const match = textFinder.matchEntireCell(true).findNext();
  
  if (match) {
    return match.getRow();
  }
  return -1;
}
