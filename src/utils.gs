/*******************************************************
 **            UTILITY AND HELPER FUNCTIONS           **
 *******************************************************/

/**
 * Validates the global SHEETCONFIG object for structural integrity.
 * 
 * It performs the following checks:
 * - Check for duplicate column names (case-insensitive)
 * - Check for duplicate column letters
 * - Check for duplicate named ranges within the spreadsheet (cross-sheet)
 * - Check for overlapping ranges within the sheet (intra-sheet)
 * 
 * Issues are logged using `Logger.log()`. This function does not throw errors.
 * 
 * @returns {boolean}
 */
function validateSHEETCONFIG()
{
  let valid = true;

  /** @type {Object<string, string>} */
  const globalRangeNameMap = {};
  let hasDuplicateNamedRanges = false;

  for (const sheetName in SHEETCONFIG) {
    Logger.log(`Checking sheet: "${sheetName}" ...`);

    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`- Sheet does not exist.`);
      continue;
    }

    const columnConfig = SHEETCONFIG[sheetName]?.layout?.columns;
    const namedRangeConfig = SHEETCONFIG[sheetName]?.layout?.namedRanges;

    if (!columnConfig) {
      Logger.log(`- Missing "columns" in sheet's configuration.`);
      valid = false;
      continue;
    }

    if (!namedRangeConfig) {
      Logger.log(`- Missing "namedRanges" in sheet's configuration.`);
      valid = false;
      continue;
    }

    const seenNames = new Set();
    const seenLetters = {};
    let hasDuplicateNames = false;
    let hasDuplicateLetters = false;

    for (const colLetter in columnConfig) {
      const colDef = columnConfig[colLetter];
      const colName = colDef.name;
      const upperLetter = colLetter.toUpperCase();

      // Check for duplicate column names (case-insensitive)
      const normalizedColName = colName.toLowerCase();
      if (seenNames.has(normalizedColName)) {
        Logger.log(`- Duplicate column name "${colName}" (case-insensitive).`);
        hasDuplicateNames = true;
      } else {
        seenNames.add(normalizedColName);
      }

      // Check for duplicate column letters
      if (seenLetters[upperLetter]) {
        Logger.log(`- Duplicate column letter "${upperLetter}" in "${seenLetters[upperLetter]}" and "${colLetter}".`);
        hasDuplicateLetters = true;
      } else {
        seenLetters[upperLetter] = colLetter;
      }
    }

    if (hasDuplicateNames) valid = false;
    else Logger.log(`- No duplicate column names found.`);

    if (hasDuplicateLetters) valid = false;
    else Logger.log(`- No duplicate column letters found.`);

    let hasRangeOverlaps = false;
    const keys = Object.keys(namedRangeConfig);
    for (let i = 0; i < keys.length; i++) {
      const rangeNameA = keys[i];
      const notationA = namedRangeConfig[rangeNameA]?.notation;

      // Skip if invalid
      if (!notationA) continue;

      // Check for duplicate named ranges within the spreadsheet (cross-sheet)
      if (globalRangeNameMap[rangeNameA]) {
        const prev = globalRangeNameMap[rangeNameA];
        Logger.log(`- Cross-sheet conflict: rangeName "${rangeNameA}" already defined in sheet "${prev}".`);
        hasDuplicateNamedRanges = true;
      } else {
        globalRangeNameMap[rangeNameA] = sheetName;
      }

      // Check for overlapping ranges within the sheet (intra-sheet)
      for (let j = i + 1; j < keys.length; j++) {
        const rangeNameB = keys[j];
        const notationB = namedRangeConfig[rangeNameB]?.notation;

        if (!notationB) continue;

        const rangeA = sheet.getRange(notationA);
        const rangeB = sheet.getRange(notationB);

        if (isRangesOverlap(rangeA, rangeB)) {
          Logger.log(`- Overlapping range "${nameA}" (${rangeA}) with "${nameB}" (${rangeB}).`);
          hasRangeOverlaps = true;
        }
      }
    }

    if (hasRangeOverlaps) valid = false;
    else Logger.log(`- No overlapping ranges found.`);

  }

  if (hasDuplicateNamedRanges) valid = false;
  Logger.log('- No duplicate named-ranges found.');

  Logger.log(`== ${valid ? 'SUCCESS' : 'FAIL'} ==`);
  return valid;
}

/**
 * Deep freezes an object, making it read-only (including nested objects).
 * 
 * @param {object} obj - The object to freeze.
 * @returns {object} The deeply frozen object.
 */
function readOnlyObject(obj)
{
  Object.getOwnPropertyNames(obj).forEach((prop) => {
    const value = obj[prop];
    if (typeof value === 'object' && value !== null) {
      readOnlyObject(value); // recursively freeze nested objects
    }
  });

  return Object.freeze(obj);
}

/**
 * Shows an alert using the UI if available, otherwise logs the message.
 *
 * @param {string} message - The message to display or log.
 * @returns {void}
 */
function showAlert(message)
{
  if (UI) {
    UI.alert(message);
  } else {
    Logger.log(message);
  }
}

/**
 * Checks if two Google Sheets ranges overlap.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} a 
 * @param {GoogleAppsScript.Spreadsheet.Range} b 
 * @returns {boolean}
 */
function isRangesOverlap(a, b)
{
  const aRowStart = a.getRow();
  const aRowEnd = aRowStart + a.getNumRows() - 1;
  const aColStart = a.getColumn();
  const aColEnd = aColStart + a.getNumColumns() - 1;

  const bRowStart = b.getRow();
  const bRowEnd = bRowStart + b.getNumRows() - 1;
  const bColStart = b.getColumn();
  const bColEnd = bColStart + b.getNumColumns() - 1;

  const rowsOverlap = aRowStart <= bRowEnd && bRowStart <= aRowEnd;
  const colsOverlap = aColStart <= bColEnd && bColStart <= aColEnd;

  return rowsOverlap && colsOverlap;
}

/**
 * Checks if a given range refers to a single cell.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to check.
 * @returns {boolean} True if the range is a single cell, false otherwise.
 */
function isSingleCell(range)
{
  return range.getNumRows() === 1 && range.getNumColumns() === 1;
}

/**
 * Applies data validation rules to columns and named ranges in the SmartSheet.
 *
 * @param {SmartSheet} smartSheet - The SmartSheet instance containing validation rule definitions.
 * @returns {void}
 */
function applyValidationRules(smartSheet)
{
  const sheet = smartSheet.sheet;

  const colLetters = smartSheet.getColumnLetters();
  for (const colName in colLetters) {
    const rule = smartSheet.getColumnValidationRule(colName);
    if (rule) {
      const colLetter = colLetters[colName];
      const headerRowCount = smartSheet.getHeaderRowCount();
      const range = sheet.getRange(`${colLetter}${headerRowCount + 1}:${colLetter}`);
      range.clearDataValidations();
      range.setDataValidation(rule);
    }
  }

  const namedRanges = sheet.getNamedRanges();
  for (const namedRange of namedRanges) {
    const rangeName = namedRange.getName();
    const rule = smartSheet.getRangeValidationRule(rangeName);
    if (rule) {
      const range = namedRange.getRange();
      range.clearDataValidations();
      range.setDataValidation(rule);
    }
  }
}

/**
 * Applies conditional formatting rules to columns and named ranges in the SmartSheet.
 *
 * @param {SmartSheet} smartSheet - The SmartSheet instance containing formatting rule definitions.
 * @returns {void}
 */
function applyFormattingRules(smartSheet)
{
  const sheet = smartSheet.sheet;
  const allRules = [];

  const colLetters = smartSheet.getColumnLetters();
  const headerRows = smartSheet.getHeaderRowCount();

  // Process formatting rules for each column
  for (const colName in colLetters) {
    const getBuilders = smartSheet.getColumnFormattingRule(colName);
    if (typeof getBuilders === 'function') {
      const builders = getBuilders(sheet, smartSheet) || [];
      const colLetter = colLetters[colName];
      const range = sheet.getRange(`${colLetter}${headerRows + 1}:${colLetter}`);
      for (const builder of builders) {
        allRules.push(builder.setRanges([range]).build());
      }
    }
  }

  // Process formatting rules for each named range
  const namedRanges = sheet.getNamedRanges();
  for (const namedRange of namedRanges) {
    const rangeName = namedRange.getName();
    const getBuilders = smartSheet.getRangeFormattingRule(rangeName);
    if (typeof getBuilders === 'function') {
      const builders = getBuilders(sheet, smartSheet) || [];
      const range = namedRange.getRange();
      for (const builder of builders) {
        allRules.push(builder.setRanges([range]).build());
      }
    }
  }

  // Apply all collected formatting rules at once
  sheet.setConditionalFormatRules(allRules);
}

/**
 * Defines named ranges in the spreadsheet for the current sheet,
 * based on the namedRanges config.
 *
 * @param {SmartSheet} smartSheet
 * @returns {void}
 */
function setNamedRanges(smartSheet)
{
  const sheet = smartSheet.sheet;
  const ss = sheet.getParent();

  const ranges = smartSheet.getNamedRangeNotations();
  for (const name in ranges) {
    const range = sheet.getRange(ranges[name]);
    ss.setNamedRange(name, range);
  }
}

/**
 * Applies calculated formulas to columns in the SmartSheet.
 *
 * @param {SmartSheet} smartSheet - The SmartSheet instance with calculated column definitions.
 * @returns {void}
 */
function applyCalculatedColumns(smartSheet)
{
  const sheet = smartSheet.sheet;
  const headerRowCount = smartSheet.getHeaderRowCount();
  const colLetters = smartSheet.getColumnLetters();

  const calculatedColumns = smartSheet.getCalculatedColumns();
  for (const column in calculatedColumns) {
    const { name, formula, lock } = calculatedColumns[column]; // formula always exist
    const expression = convertExpression(formula, colLetters);
    const formulas = expressionRange(expression, headerRowCount + 1, sheet.getMaxRows());
    const range = sheet.getRange(`${column}${headerRowCount + 1}:${column}`); // usually the whole column (exclude headers)
    range.setFormulas(formulas);
    if (lock) lockRange(range, name);
  }
}

/**
 * Applies calculated formulas to named ranges in the given SmartSheet based on its configured expressions.
 *
 * @param {SmartSheet} smartSheet - An instance of SmartSheet that contains named ranges and column mappings.
 * @returns {void}
 */
function applyCalculatedNamedRanges(smartSheet)
{
  const rangeNotations = smartSheet.getNamedRangeNotations();

  const calculatedRanges = smartSheet.getCalculatedNamedRanges();
  for (const notation in calculatedRanges) {
    const { name, args, formula, lock } = calculatedRanges[notation]; // formula always exist
    const range = smartSheet.getNamedRange(name);
    const expression = isFunction(formula)
      ? `${formula.name}(${args.map(arg => smartSheet.getNamedRangeNotations()[arg]).join(', ')})`
      : convertExpression(formula, rangeNotations, false);
    if (isSingleCell(range)) range.setFormula(`=${expression}`); // move this check to SHEETCONFIG validation
    if (lock) lockRange(range, name);
  }
}

/**
 * Converts a formula string by replacing variable names (e.g., $score) with their mapped column letters.
 *
 * @param {string} str - The formula string containing placeholders like $columnName.
 * @param {Object.<string, string>} map - An object mapping column names to column letters (e.g., { score: 'E' }).
 * @param {boolean} [withSign=true] - If true, retains the `$` prefix (e.g., $E); otherwise, removes it.
 * @returns {string} - The converted formula string with mapped column letters.
 */
function convertExpression(str, map, withSign = true)
{
  const regex = /\$([a-zA-Z0-9_]+)/g;
  return str.replace(regex, (_, key) => {
    return withSign ? `$${map[key]}` : `${map[key]}`;
  });
}

/**
 * Expands an expression with `$<columnLetter>` placeholders into row-specific formulas.
 *
 * @param {string} expression - The expression containing placeholders like $E.
 * @param {number} startIndex - The starting row number (inclusive).
 * @param {number} lastIndex - The ending row number (inclusive).
 * @returns {string[][]} - A 2D array where each subarray contains one formula string with row numbers applied.
 */
function expressionRange(expression, startIndex, lastIndex)
{
  const result = [];

  for (let row = startIndex; row <= lastIndex; row++) {
    const regex = /\$([A-Z]+)/g;
    const rowExpr = expression.replace(regex, (_, colLetter) => {
      return `${colLetter}${row}`;
    });
    result.push([`=${rowExpr}`]);
  }

  return result;
}

/**
 * Protects a given range in the active spreadsheet and allows specific editors.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to protect.
 * @param {string} name - The name of the range.
 * @returns {void}
 */
function lockRange(range, name)
{
  const sheet = range.getSheet();
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  let protected;
  for (const protection of protections) {
    const pRange = protection.getRange();
    if (pRange.getA1Notation() === range.getA1Notation()) {
      protected = protection;
    }
  }
  protected = protected || range.protect();

  protected.setDescription(name);

  const protectors = new Set(
    editors.filter(e => e.protector).map(e => e.email)
  );

  const currentProtectors = new Set(
    protected.getEditors().map(e => e.getEmail())
  );

  // Remove editors not in the allowed list
  const emailsToRemove = [...currentProtectors].filter(email => !protectors.has(email));
  if (emailsToRemove.length > 0) {
    protected.removeEditors(emailsToRemove);
  }

  // Add editors not already present
  const emailsToAdd = [...protectors].filter(email => !currentProtectors.has(email));
  if (emailsToAdd.length > 0) {
    protected.addEditors(emailsToAdd);
  }
}

/**
 * Reset a spreadsheet to a clean state.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The spreadsheet to clear.
 * @returns {void}
 */
function clearAll(ss)
{
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const numRows = sheet.getMaxRows();
    const numCols = sheet.getMaxColumns();
    const fullRange = sheet.getRange(1, 1, numRows, numCols);

    // Clear conditional formatting rules
    sheet.setConditionalFormatRules([]);

    // Clear data validations
    fullRange.clearDataValidations();

    // Clear manual formatting (font, background, borders, etc.)
    fullRange.clearFormat();

    // Remove range and sheet protections
    const protections = [
      ...sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE),
      ...sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    ];
    protections.forEach(protection => protection.remove());
  });

  // Clear all named ranges
  const namedRanges = ss.getNamedRanges();
  namedRanges.forEach(namedRange => namedRange.remove());
}

/**
 * Verify that the variable is a function.
 * Exclude AsyncFunction and GeneratorFunction.
 * 
 * @param {function} func - The variable to check.
 * @returns {boolean}
 */
function isFunction(func)
{
  return Object.prototype.toString.call(func) === '[object Function]';
}
