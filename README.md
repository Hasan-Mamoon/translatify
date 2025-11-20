# Translatify

A helper tool designed to support collaboration in multinational organizations. Translatify allows users to insert translations of specified columns directly into their spreadsheets, creating new translated columns next to the original ones (currently limited to Google Sheets).

## Features

- **Cell-based column selection**: Specify exact cell references (e.g., `M2`) to translate columns starting from any row
- **Multiple column translation**: Translate multiple columns in a single operation
- **Automatic column insertion**: Translated columns are automatically inserted immediately to the right of the original columns
- **Robust error handling**: Safely handles images, formulas, errors, and unsupported data types
- **Header translation**: Automatically translates column headers
- **Smart processing**: Processes columns from right to left to prevent index shifting issues

## Installation

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any default code in the editor
4. Copy and paste the contents of `google-sheets/translatify.gs` into the editor
5. Save the project (Ctrl+S or Cmd+S)
6. Reload your Google Sheet
7. You should now see a new **Translatify** menu in the menu bar
8. When running for the first time, you will be asked for authorization:
   - On the Google consent screen, click **Advanced** (bottom left)
   - Click **Go to untitled project (unsafe)**
   - On the new screen, scroll down and click **Continue** to grant access

## Usage

### Basic Workflow

1. Click **Translatify → Translate columns…** from the menu
2. Enter the **target language code** (e.g., `ja` for Japanese, `es` for Spanish, `fr` for French)
3. Enter **cell references** for the columns you want to translate, separated by commas (e.g., `B2, M2`)
   - Each cell reference specifies the starting cell of a column to translate
   - Format: `[Column Letter][Row Number]` (e.g., `M2` means column M starting from row 2)
4. Click **OK** and wait for the translation to complete

### Example

**Before:**
```
| id | testcase | steps | explanation |
|----|----------|-------|-------------|
| 1  | Login    | Click | User login  |
```

**After** (translating `testcase` and `steps` columns to Japanese with references `B2, C2`):
```
| id | testcase | テストケース | steps | 手順 | explanation |
|----|----------|-------------|-------|------|-------------|
| 1  | Login    | ログイン     | Click | クリック | User login  |
```

## How It Works

### Translation Engine

Translatify uses Google's `LanguageApp.translate()` API, which:
- Automatically detects the source language (when source language is empty string `''`)
- Supports 100+ languages via ISO 639-1 language codes
- Translates text content in cells
- Handles common text, numbers, and formula results

### Cell Reference Parsing

- Cell references are parsed using regex pattern: `^([A-Za-z]+)(\d+)$`
- Column letters are converted to 1-based column indices (A=1, Z=26, AA=27, etc.)
- The row number determines where translation starts in that column

### Error Handling

The tool includes robust error handling for various edge cases:

1. **Empty/Null Values**: Empty cells remain empty in translated columns
2. **Non-Text Types**: 
   - Booleans, objects, and functions are preserved as-is (not translated)
   - Images and drawings are skipped (original value preserved)
3. **Long Text**: Text longer than 5000 characters is truncated to prevent quota issues
4. **Translation Failures**: If translation fails for any cell, the original value is preserved to prevent data loss
5. **Invalid References**: Invalid cell references are filtered out with user notification

### Processing Order

Columns are processed from **right to left** to prevent column index shifting when new columns are inserted. This ensures accurate placement of translated columns.

## Supported Language Codes

Common language codes (ISO 639-1):

- `ja` - Japanese
- `es` - Spanish
- `fr` - French
- `de` - German
- `zh` - Chinese
- `ar` - Arabic
- `pt` - Portuguese
- `ru` - Russian
- `it` - Italian
- `ko` - Korean
- `hi` - Hindi
- `nl` - Dutch
- `pl` - Polish
- `tr` - Turkish
- `sv` - Swedish
- `da` - Danish
- `fi` - Finnish
- `no` - Norwegian
- `th` - Thai
- `vi` - Vietnamese

For a complete list, see [Google Translate supported languages](https://cloud.google.com/translate/docs/languages).

## Limitations

- **Translation Quotas**: Google Apps Script has daily quotas for translation API calls. Very large sheets may hit these limits
- **Text Only**: Images, charts, and complex objects in cells are not translated (preserved as-is)
- **Formula Results**: Only the displayed value of formulas is translated, not the formula itself
- **Length Limits**: Text longer than 5000 characters per cell is truncated
- **Single Sheet**: Currently processes only the active sheet

## Troubleshooting

### "No valid cell references" error
- Ensure cell references are in the correct format: `[Letter][Number]` (e.g., `B2`, `M10`)
- Check that references don't contain spaces (use `B2, M2` not `B 2, M 2`)

### Translation not working
- Verify the language code is correct (use ISO 639-1 format)
- Check that cells contain translatable text (not just images or errors)
- Ensure you have sufficient Google Apps Script quota remaining

### Columns in wrong position
- The tool inserts columns immediately to the right of the original
- If you need different positioning, manually rearrange columns after translation
