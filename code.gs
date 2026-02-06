/**
 * @fileoverview Online Genko Manuscript Paper - Server-side Logic
 * Handles interactions with Google Sheet database.
 * Features:
 * - Logical deletion (Archiving)
 * - Concurrency control (LockService)
 * - Performance optimization (Limit retrieval)
 */

// --- Global Constants ---
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_NAME = '作文データ';
const SHEET = SPREADSHEET.getSheetByName(SHEET_NAME);

/**
 * Serves the main HTML page.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('オンライン原稿用紙')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setFaviconUrl('https://drive.google.com/uc?id=1EsaLbGPFc9WixYhJ5sPynTIBZpxzSsfK&.png');
}

/**
 * Saves a new draft or updates an existing one with Concurrency Control.
 */
function saveOrUpdateDraft(draftData) {
  // Lock to prevent concurrent writes
  const lock = LockService.getScriptLock();
  // Wait up to 10 seconds for other processes to finish
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return { status: 'error', message: '他の人が保存中です。もう一度試してください。' };
  }

  try {
    const now = new Date();
    if (!SHEET) throw new Error('シート「' + SHEET_NAME + '」が見つかりません。');

    // Fetch current data inside the lock to ensure freshness
    const lastRow = SHEET.getLastRow();
    
    if (draftData.id) {
      // --- Update existing draft ---
      // Check data range only if data exists
      if (lastRow > 1) {
        const ids = SHEET.getRange(2, 1, lastRow - 1, 1).getValues().flat(); // Get only IDs for speed
        const index = ids.indexOf(draftData.id);

        if (index !== -1) {
          const rowIndex = index + 2; // 1-based index + header
          
          // Preserve original Creation Date (Column F / index 6)
          const creationDate = SHEET.getRange(rowIndex, 6).getValue();

          SHEET.getRange(rowIndex, 1, 1, 8).setValues([[
            draftData.id,
            draftData.title,
            draftData.class,
            draftData.name,
            draftData.content,
            creationDate,
            now, // Update Date
            ''   // Clear DeletedAt (Un-archive if needed)
          ]]);
          return { status: 'success', message: '作文を更新しました。', id: draftData.id };
        }
      }
    }
    
    // --- Save new draft ---
    const newId = Utilities.getUuid();
    SHEET.appendRow([
      newId,
      draftData.title,
      draftData.class,
      draftData.name,
      draftData.content,
      now, // Creation Date
      now, // Update Date
      ''   // DeletedAt
    ]);
    return { status: 'success', message: '下書きを保存しました。', id: newId };

  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: '保存中にエラーが発生しました: ' + e.message };
  } finally {
    lock.releaseLock(); // Always release the lock
  }
}

/**
 * Retrieves a list of active drafts.
 * Limits to the latest 50 items for performance.
 */
function getDraftList() {
  try {
    if (!SHEET || SHEET.getLastRow() < 2) return [];
    
    // Fetch all data (needed for filtering and sorting correctly)
    // Note: For very large datasets (10k+), we might need more optimized query logic,
    // but fetching all is usually fine for text data within GAS limits.
    const data = SHEET.getRange(2, 1, SHEET.getLastRow() - 1, 8).getValues();
    
    // 1. Filter out archived (deleted) drafts
    const activeDrafts = data.filter(row => row[7] === '');

    // 2. Map to lightweight objects
    const drafts = activeDrafts.map(row => ({
      id: row[0],
      title: row[1],
      name: row[3],
      updatedAt: new Date(row[6]) // Date object for sorting
    }));
    
    // 3. Sort by update date (Newest first)
    drafts.sort((a, b) => b.updatedAt - a.updatedAt);

    // 4. Limit to latest 50 items
    const limitedDrafts = drafts.slice(0, 50);

    // 5. Format date string for client
    return limitedDrafts.map(d => ({
      ...d,
      updatedAt: d.updatedAt.toLocaleString('ja-JP', { 
        year: 'numeric', month: '2-digit', day: '2-digit', 
        hour: '2-digit', minute: '2-digit' 
      })
    }));

  } catch (e) {
    Logger.log(e);
    return []; 
  }
}

/**
 * Loads the full content of a draft.
 */
function loadDraft(id) {
  try {
    const data = SHEET.getDataRange().getValues();
    // Simple linear search is robust enough for small-medium datasets
    const row = data.find((r, i) => i > 0 && r[0] == id);
    
    if (row) {
      if (row[7] !== '') {
        return { status: 'error', message: 'この作文は削除されています。' };
      }
      return {
        status: 'success',
        data: {
          id: row[0], title: row[1], class: row[2],
          name: row[3], content: row[4]
        }
      };
    }
    return { status: 'error', message: '指定された作文が見つかりませんでした。' };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: '読み込み中にエラーが発生しました: ' + e.message };
  }
}

/**
 * Logically deletes a draft (Archives it).
 * Uses LockService to prevent conflicts during deletion.
 */
function deleteDraft(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); // Wait up to 5 sec
  } catch (e) {
    return { status: 'error', message: '他の人が操作中です。' };
  }

  try {
    const data = SHEET.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex !== -1) {
      // Set delete timestamp in column H (8)
      SHEET.getRange(rowIndex, 8).setValue(new Date());
      return { status: 'success', message: '作文をゴミ箱（アーカイブ）へ移動しました。' };
    }
    return { status: 'error', message: '削除対象の作文が見つかりませんでした。' };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: '削除処理中にエラーが発生しました: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}
