function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Translatify')
    .addItem('Translate columns…', 'showTranslatePrompt')
    .addToUi();
}

function showTranslatePrompt() {
  const ui = SpreadsheetApp.getUi();

  const langResp = ui.prompt(
    'Target language code',
    'Enter a language code (e.g. "ja" for Japanese):',
    ui.ButtonSet.OK_CANCEL
  );
  if (langResp.getSelectedButton() !== ui.Button.OK) return;
  const targetLang = langResp.getResponseText().trim();
  if (!targetLang) {
    ui.alert('No language code provided.');
    return;
  }

  const refResp = ui.prompt(
    'Column start cells',
    'Enter cell references for the columns to translate (e.g. "B2, M2"):',
    ui.ButtonSet.OK_CANCEL
  );
  if (refResp.getSelectedButton() !== ui.Button.OK) return;
  const refs = refResp
    .getResponseText()
    .split(',')
    .map(r => r.trim())
    .filter(Boolean);

  if (!refs.length) {
    ui.alert('No cell references provided.');
    return;
  }

  translateColumnsFromCells(refs, targetLang);
}

function translateColumnsFromCells(cellRefs, targetLang) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Not enough data to translate.');
    return;
  }

  const parsedColumns = cellRefs
    .map(ref => parseCellReference(ref))
    .filter(entry => entry !== null);

  if (!parsedColumns.length) {
    SpreadsheetApp.getUi().alert('No valid cell references.');
    return;
  }

  // Process from right to left to avoid column index shifts
  parsedColumns.sort((a, b) => b.col - a.col);

  parsedColumns.forEach(({ col, rowStart, label }) => {
    const numRows = sheet.getLastRow() - rowStart + 1;
    if (numRows <= 0) return;

    const range = sheet.getRange(rowStart, col, numRows, 1);
    const values = range.getValues();

    const translated = values.map(row => {
      const original = row[0];

      // 1. Treat empty / null / error-like values as blank
      if (original === '' || original === null || original === undefined) {
        return [''];
      }

      // 2. Skip obvious non-text types (e.g. booleans, objects)
      const valueType = typeof original;
      if (valueType === 'boolean' || valueType === 'object' || valueType === 'function') {
        // Keep as-is (or choose to blank it out instead)
        return [original];
      }

      // 3. Convert numbers and other primitives to string
      let text = String(original).trim();
      if (!text) {
        return [''];
      }

      // 4. Optional: limit very long strings to avoid quota / performance issues
      const MAX_LEN = 5000; // adjust as needed
      if (text.length > MAX_LEN) {
        // You can choose: cut, skip, or keep original
        text = text.slice(0, MAX_LEN);
      }

      // 5. Translate with robust error handling
      try {
        const t = LanguageApp.translate(text, '', targetLang);
        return [t];
      } catch (e) {
        // Log the error for developers and keep original to avoid data loss
        console.warn('Translation error for value:', original, 'Error:', e);
        return [original];
      }
    });

    sheet.insertColumnAfter(col);
    const newCol = col + 1;

    let translatedHeader = label;
    try {
      translatedHeader = LanguageApp.translate(label, '', targetLang);
    } catch (e) {}

    sheet.getRange(rowStart - 1 > 0 ? rowStart - 1 : rowStart, newCol).setValue(translatedHeader);
    sheet.getRange(rowStart, newCol, numRows, 1).setValues(translated);
  });

  SpreadsheetApp.getUi().alert(
    `Translation complete.\nColumns: ${cellRefs.join(', ')}\nLanguage: ${targetLang}`
  );
}

function parseCellReference(ref) {
  const match = /^([A-Za-z]+)(\d+)$/.exec(ref);
  if (!match) return null;
  const [, colLetters, rowNumber] = match;
  const colIndex = lettersToColumn(colLetters.toUpperCase());
  return { col: colIndex, rowStart: parseInt(rowNumber, 10), label: ref.toUpperCase() };
}

function lettersToColumn(letters) {
  return letters.split('').reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0);
}
