/**
 * Archive Automation Script
 *
 * This script automates the process of analyzing archival PDFs from a Google
 * Drive folder, extracting metadata using the Gemini API, and recording the
 * results in a Google Sheet.
 *
 * @version 1.6 - Corrected advanced Drive API call from insert() to create().
 */

// ===============================================================
// === CONFIGURATION                                         ===
// ===============================================================

// --- USER CONFIGURATION ---

// 1. The ID of the Google Drive folder containing the archival files.
//    (Found in the folder's URL: .../folders/THIS_IS_THE_ID)
const FOLDER_ID = '###';

// 2. The ID of the Google Sheet.
//    (Found in the sheet's URL: .../d/THIS_IS_THE_ID/edit)
const SPREADSHEET_ID = '###';

// 3. Gemini API Key.
//    (From Google AI Studio: https://aistudio.google.com/app/apikey)
const API_KEY = '###';

// 4. The Gemini model to use for analysis. 2.5-flash is recommended.
const GEMINI_MODEL = 'gemini-2.5-flash-preview-05-20';

// 5. Sheet tab name where archival file data should be updated
const SHEET_TAB_NAME = '###'

// --- END OF USER CONFIGURATION ---


// ===============================================================
// === SPREADSHEET MENU (FOR MANUAL TESTING)                 ===
// ===============================================================

/**
 * Creates a custom menu in the spreadsheet to allow for manual script execution.
 * This function only runs automatically when the script is bound to the spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Archival Tools')
      .addItem('Process New Files Manually', 'processNewFiles')
      .addToUi();
}


// ===============================================================
// === CORE FUNCTIONS                                        ===
// ===============================================================

/**
 * Main function to be triggered.
 * It finds unprocessed PDF files, processes them, and logs the results.
 */
function processNewFiles() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByType(MimeType.PDF);
  const processedPdfLinks = getProcessedPdfLinks_();
  let filesProcessed = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    // Skip temporary files created by our script
    if (fileName.startsWith('[TEMP]')) {
      Logger.log(`Skipping temporary file: ${fileName}`);
      continue;
    }

    const pdfLink = file.getUrl();
    if (!processedPdfLinks.has(pdfLink)) {
      filesProcessed++;
      Logger.log(`--- Starting processing for new file: ${fileName} (ID: ${file.getId()}) ---`);

      try {
        Logger.log('Step 1: Starting text extraction...');
        const textContent = extractTextFromPdf_(file);

        if (!textContent) {
          Logger.log(`Error: Could not extract text from: ${fileName}. Skipping this file.`);
          continue;
        }

        if (textContent.length < 50) {
          Logger.log(`Warning: Very little text extracted from: ${fileName} (${textContent.length} chars). Skipping this file.`);
          continue;
        }

        Logger.log(`Step 2: Text extracted successfully (${textContent.length} chars). First 100 chars: ${textContent.substring(0, 100)}`);
        Logger.log('Step 3: Calling Gemini API...');
        const aiData = getAiAnalysis_(textContent);

        if (aiData) {
          Logger.log('Step 4: AI analysis successful. Received data object.');
          Logger.log(JSON.stringify(aiData, null, 2));
          Logger.log('Step 5: Logging data to sheet...');
          logDataToSheet_(file, aiData);
          Logger.log(`--- Successfully processed and logged: ${fileName} ---`);
        } else {
          Logger.log(`Error: Failed to get AI analysis for: ${fileName}. Skipping this file.`);
        }

      } catch (e) {
        Logger.log(`Critical Error during processing for file ${fileName}: ${e.toString()}`);
        Logger.log(`Stack Trace: ${e.stack}`);
      }
    }
  }

  Logger.log(`Processing run complete. Attempted to process ${filesProcessed} new file(s).`);
  SpreadsheetApp.getUi().alert(`Processing complete. Checked for new files and attempted to process ${filesProcessed} file(s). Check logs for details.`);
}


/**
 * Extracts text from a PDF file by converting it to a temporary Google Doc.
 * @param {File} file The PDF file object from Drive.
 * @return {string|null} The extracted text content, or null on failure.
 * @private
 */
function extractTextFromPdf_(file) {
  let tempDocId = null;

  try {
    Logger.log(`Starting PDF text extraction for: ${file.getName()}`);

    const blob = file.getBlob();

    if (blob.getContentType() !== 'application/pdf') {
      Logger.log(`File is not a PDF. Content type: ${blob.getContentType()}`);
      return null;
    }

    // Create the temporary Google Doc with OCR - using different approach
    const tempDocFile = Drive.Files.create(
      {
        name: `[TEMP] OCR - ${file.getName()}`
      },
      blob,
      {
        ocr: true,
        ocrLanguage: 'en',
        convert: true  // This is the key parameter that was missing!
      }
    );

    if (!tempDocFile || !tempDocFile.id) {
      Logger.log('Failed to create temporary Google Doc for OCR');
      return null;
    }

    tempDocId = tempDocFile.id;
    Logger.log(`Created temporary doc with ID: ${tempDocId}`);

    // Wait for OCR processing
    Utilities.sleep(8000); // Increased wait time

    // Check what type of file was actually created
    const createdFile = DriveApp.getFileById(tempDocId);
    const mimeType = createdFile.getBlob().getContentType();
    Logger.log(`Created file MIME type: ${mimeType}`);

    let text = null;

    if (mimeType === 'application/vnd.google-apps.document') {
      // It's a Google Doc, we can use DocumentApp
      Logger.log('File converted to Google Doc, using DocumentApp...');
      const tempDoc = DocumentApp.openById(tempDocId);
      text = tempDoc.getBody().getText();

    } else {
      // Fall back to Drive API export
      Logger.log('File not converted to Google Doc, using Drive API export...');

      const exportResponse = UrlFetchApp.fetch(
        `https://www.googleapis.com/drive/v3/files/${tempDocId}/export?mimeType=text/plain`,
        {
          headers: {
            'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`
          },
          muteHttpExceptions: true
        }
      );

      if (exportResponse.getResponseCode() === 200) {
        text = exportResponse.getContentText();
      } else {
        Logger.log(`Drive API export failed: ${exportResponse.getResponseCode()} - ${exportResponse.getContentText()}`);

        // Last resort: try to create a new doc with explicit conversion
        Logger.log('Trying alternative conversion method...');
        const altResource = {
          name: `[TEMP] ALT - ${file.getName()}`,
          mimeType: 'application/vnd.google-apps.document'
        };

        const altDocFile = Drive.Files.create(altResource, blob, {ocr: true});

        if (altDocFile && altDocFile.id) {
          Utilities.sleep(5000);
          const altDoc = DocumentApp.openById(altDocFile.id);
          text = altDoc.getBody().getText();
          // Clean up the alternative doc too
          DriveApp.getFileById(altDocFile.id).setTrashed(true);
        }
      }
    }

    if (text) {
      Logger.log(`Successfully extracted ${text.length} characters`);
      return text;
    } else {
      Logger.log('No text could be extracted');
      return null;
    }

  } catch (error) {
    Logger.log(`Error in extractTextFromPdf_: ${error.toString()}`);
    Logger.log(`Error stack: ${error.stack}`);
    return null;

  } finally {
    // Clean up the temporary file
    if (tempDocId) {
      try {
        DriveApp.getFileById(tempDocId).setTrashed(true);
        Logger.log(`Cleaned up temporary doc with ID: ${tempDocId}`);
      } catch (cleanupError) {
        Logger.log(`Warning: Could not clean up temporary doc: ${cleanupError.toString()}`);
      }
    }
  }
}

/**
 * Sends text content to the Gemini API for analysis.
 * @param {string} text The text to analyze.
 * @return {Object|null} The parsed JSON object from the API, or null on failure.
 * @private
 */
function getAiAnalysis_(text) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${API_KEY}`;

  const payload = {
    contents: [{
      parts: [{
        text: `Based on the following text, extract the date, speaker, title, the primary theme, and any Tanakh or Talmud references as "book chapter:verse". Provide the output in a single JSON object with the keys "date", "speaker", "title", "theme" (as a single string), and "references" (as an array of strings). Do not include any other text, formatting, or markdown backticks. Text:\n\n---\n\n${text}`
      }]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    try {
      const jsonResponse = JSON.parse(responseBody);
      // Clean potential markdown code block fences from the response
      let content = jsonResponse.candidates[0].content.parts[0].text;
      content = content.replace(/```json/g, '').replace(/```/g, '').trim();
      return JSON.parse(content);
    } catch (e) {
      Logger.log(`Failed to parse Gemini API response: ${e.toString()}`);
      Logger.log(`Raw response body: ${responseBody}`);
      return null;
    }
  } else {
    Logger.log(`Gemini API Error - Response Code: ${responseCode}`);
    Logger.log(`Gemini API Error - Response Body: ${responseBody}`);
    return null;
  }
}

/**
 * Writes the processed data to the designated Google Sheet.
 * @param {File} file The original file object.
 * @param {Object} data The analysis data from the AI.
 * @private
 */
function logDataToSheet_(file, data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_TAB_NAME);
  sheet.appendRow([
    data.date || '',
    data.speaker || '',
    '', // Placeholder for Speaker type
    '', // Placeholder for Speech type
    '', // Placeholder for Event type
    data.title || '',
    data.theme || '',
    '', // Placeholder for Media Type
    (data.references || []).join(', '),
    file.getUrl(), // PDF Link
    '', // Placeholder for MP3 Link
    ''  // Placeholder for Vimeo Link
  ]);
}

/**
 * Retrieves a set of PDF links that have already been processed.
 * @return {Set<string>} A Set containing the URLs of processed files.
 * @private
 */
function getProcessedPdfLinks_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_TAB_NAME);
  if (sheet.getLastRow() < 2) {
    return new Set(); // Return an empty set if there are no data rows
  }
  // Assumes PDF Link is in the tenth column (J)
  const range = sheet.getRange(2, 10, sheet.getLastRow() - 1, 1);
  const values = range.getValues();
  const linkSet = new Set();
  values.forEach(row => {
    if (row[0]) {
      linkSet.add(row[0].toString());
    }
  });
  return linkSet;
}

