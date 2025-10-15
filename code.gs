/**
 * @fileoverview Online Genko Manuscript Paper - Server-side Logic
 * Handles all interactions with the Google Sheet database.
 */

// --- Global Constants ---
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_NAME = '作文データ';
const SHEET = SPREADSHEET.getSheetByName(SHEET_NAME);

/**
 * Serves the main HTML page of the web application.
 * This is the entry point for users accessing the app's URL.
 * @returns {HtmlOutput} The HTML content of the web application.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('オンライン原稿用紙')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Saves a new draft or updates an existing one in the spreadsheet.
 * Generates a unique ID for new drafts.
 * @param {object} draftData - The draft object from the client.
 * @param {string|null} draftData.id - The unique ID of the draft, or null for a new one.
 * @param {string} draftData.title - The title of the draft.
 * @param {string} draftData.class - The class information.
 * @param {string} draftData.name - The author's name.
 * @param {string} draftData.content - The main body of the composition.
 * @returns {object} A result object with status, message, and the draft's ID.
 */
function saveOrUpdateDraft(draftData) {
  try {
    const now = new Date();
    if (draftData.id) {
      // --- Update existing draft ---
      const data = SHEET.getDataRange().getValues();
      const header = data.shift();
      const idColIndex = header.indexOf('ID');

      let rowIndex = -1;
      for (let i = 0; i < data.length; i++) {
        if (data[i][idColIndex] == draftData.id) {
          rowIndex = i + 2; // +1 for header offset, +1 for 1-based index
          break;
        }
      }

      if (rowIndex !== -1) {
        SHEET.getRange(rowIndex, 1, 1, 7).setValues([[
          draftData.id,
          draftData.title,
          draftData.class,
          draftData.name,
          draftData.content,
          SHEET.getRange(rowIndex, 6).getValue(), // Keep original creation date
          now
        ]]);
        return { status: 'success', message: '作文を更新しました。', id: draftData.id };
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
      now  // Update Date
    ]);
    return { status: 'success', message: '下書きを保存しました。', id: newId };

  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: '保存中にエラーが発生しました: ' + e.message };
  }
}

/**
 * Retrieves a list of all saved drafts.
 * @returns {Array<object>} An array of draft objects, sorted by most recently updated.
 */
function getDraftList() {
  try {
    if (SHEET.getLastRow() < 2) return [];
    
    const data = SHEET.getRange(2, 1, SHEET.getLastRow() - 1, 7).getValues();
    const drafts = data.map(row => ({
      id: row[0],
      title: row[1],
      name: row[3],
      updatedAt: new Date(row[6]).toLocaleString('ja-JP')
    }));
    
    // Sort by update date, descending
    return drafts.sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
  } catch (e) {
    Logger.log(e);
    return []; // Return empty on error
  }
}

/**
 * Loads the full content of a specific draft by its ID.
 * @param {string} id - The unique ID of the draft to load.
 * @returns {object} A result object containing status and draft data, or an error message.
 */
function loadDraft(id) {
  try {
    const data = SHEET.getDataRange().getValues();
    const header = data.shift();
    const idColIndex = header.indexOf('ID');
    
    const row = data.find(r => r[idColIndex] == id);
    
    if (row) {
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
 * Deletes a draft from the spreadsheet by its ID.
 * @param {string} id - The unique ID of the draft to delete.
 * @returns {object} A result object with status and a message.
 */
function deleteDraft(id) {
  try {
    const data = SHEET.getDataRange().getValues();
    const header = data.shift();
    const idColIndex = header.indexOf('ID');

    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) {
        if (data[i][idColIndex] == id) {
          rowIndex = i + 2; // +1 for header, +1 for 1-based index
          break;
        }
    }

    if (rowIndex !== -1) {
      SHEET.deleteRow(rowIndex);
      return { status: 'success', message: '作文を削除しました。' };
    }
    return { status: 'error', message: '削除対象の作文が見つかりませんでした。' };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: '削除中にエラーが発生しました: ' + e.message };
  }
}
