
/**
 * A utility class for handling structured data access in a Google Sheets sheet.
 *
 * This class wraps a GoogleAppsScript `Sheet` object and provides a structured way
 * to access the Sheet using a centralized config.
 *
 * It relies on a global `SHEETCONFIG` object.
 */
class SmartSheet
{
  /**
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to wrap.
   */
  constructor(sheet)
  {
    this.sheet = sheet;
    this.sheetName = sheet.getName();

    // Memoize the lookups
    this._calculatedColumns = null;
    this._calculatedNamedRanges = null;
    this._columnLetters = null;
    this._columnTypes = null;
    this._columnNumbers = null;
    this._namedRangeNotations = null;
  }

  /**
   * Get the number of header rows configured for the sheet.
   * Defaults to 0 if not specified in the config.
   * 
   * @returns {number} Number of header rows.
   */
  getHeaderRowCount()
  {
    return SHEETCONFIG[this.sheetName]?.layout?.headerRows ?? 0;
  }

  /**
   * Retrieves the column configuration for the current sheet from the global SHEETCONFIG.
   * 
   * @returns {Object<string, { name: string, type: string }>}
   *          An object mapping column letters to their configuration.
   */
  getColumnConfig()
  {
    return SHEETCONFIG[this.sheetName]?.layout?.columns ?? {};
  }

  /**
   * Returns the named range configuration for the current sheet from the global SHEETCONFIG.
   *
   * @returns {Object<string, { name: string, type: string }>}
   *          An object mapping range notations to their configuration.
   */
  getNamedRangeConfig()
  {
    return SHEETCONFIG[this.sheetName]?.layout?.namedRanges ?? {};
  }

  /**
   * Retrieves column configurations that include a formula.
   *
   * @returns {Object.<string, { name: string, formula: string, lock?: boolean }>}
   *          An object mapping column letters to their config objects containing formulas.
   */
  getCalculatedColumns()
  {
    if (this._calculatedColumns !== null) return this._calculatedColumns;

    const colConfig = this.getColumnConfig();
    this._calculatedColumns = {};
    for (const colLetter in colConfig) {
      if (colConfig[colLetter].hasOwnProperty('formula')) {
        this._calculatedColumns[colLetter] = colConfig[colLetter];
      }
    }
    return this._calculatedColumns;
  }

  /**
   * Retrieves named ranges that have an associated formula in their configuration.
   *
   * @returns {Object.<string, { name: string, formula: string, lock?: boolean }>}
   *          An object mapping range notations to their config objects containing formulas.
   */
  getCalculatedNamedRanges()
  {
    if (this._calculatedNamedRanges !== null) return this._calculatedNamedRanges;

    const rangeConfig = this.getNamedRangeConfig();
    this._calculatedNamedRanges = {};
    for (const notation in rangeConfig) {
      if (rangeConfig[notation].hasOwnProperty('formula')) {
        this._calculatedNamedRanges[notation] = rangeConfig[notation];
      }
    }
    return this._calculatedNamedRanges;
  }

  /**
   * Get a map of column names to column letters.
   * 
   * @returns {Object<string, string>} column-name → column-letter (e.g., "A", "B").
   */
  getColumnLetters()
  {
    if (this._columnLetters !== null) return this._columnLetters;

    const colConfig = this.getColumnConfig();
    this._columnLetters = {};
    for (const colLetter in colConfig) {
      const { name } = colConfig[colLetter];
      this._columnLetters[name] = colLetter.toUpperCase();
    }
    return this._columnLetters;
  }

  /**
   * Get a map of column names to data types.
   * 
   * @returns {Object<string, string>} column-name → data-type (e.g., "string", "number").
   */
  getColumnTypes()
  {
    if (this._columnTypes !== null) return this._columnTypes;

    const colConfig = this.getColumnConfig();
    this._columnTypes = {};
    for (const colLetter in colConfig) {
      const { name, type } = colConfig[colLetter];
      this._columnTypes[name] = type;
    }
    return this._columnTypes;
  }

  /**
   * Get a map of column names to column indexes (1-based).
   * 
   * @returns {Object<string, number>} column-name → column-index (e.g., 1, 2).
   */
  getColumnNumbers()
  {
    if (this._columnNumbers !== null) return this._columnNumbers;

    const colLetters = this.getColumnLetters();
    this._columnNumbers = {};
    for (const name in colLetters) {
      const letter = colLetters[name];
      this._columnNumbers[name] = letter
        .split('')
        .reduce((total, char) => total * 26 + (char.charCodeAt(0) - 64), 0);
    }
    return this._columnNumbers;
  }

  /**
   * Returns a mapping of named range names to their A1-style notations (in uppercase).
   *
   * @returns {Object.<string, string>} An object mapping named range names to notations.
   */
  getNamedRangeNotations()
  {
    if (this._namedRangeNotations !== null) return this._namedRangeNotations;

    const namedRangeConfig = this.getNamedRangeConfig();
    this._namedRangeNotations = {};
    for (const notation in namedRangeConfig) {
      const { name } = namedRangeConfig[notation];
      this._namedRangeNotations[name] = notation.toUpperCase();
    }
    return this._namedRangeNotations;
  }

  /**
   * Returns all values from the specified column name, excluding header rows.
   *
   * @param {string} colName - The name of the column as defined in config.
   * @returns {any[]} Array of values from that column.
   */
  getColumnValues(colName)
  {
    const headerRows = this.getHeaderRowCount();
    const lastRow = this.sheet.getLastRow();
    const col = this.getColumnNumbers()[colName];

    if (!col) throw new Error(`Unknown column name: ${colName}`);

    const range = this.sheet.getRange(headerRows + 1, col, lastRow - headerRows, 1);
    return range.getValues().flat();
  }

  /**
   * Returns all values in the specified row as a flat array.
   * Returns null if the row is within the configured header rows.
   * Throws an error if the sheet has no data.
   * 
   * @param {number} rowNumber - Actual 1-based row number in the sheet.
   * @returns {any[] | null} Array of values from the row, or null if within header.
   * @throws {Error} If the sheet has no data.
   */
  getRowValues(rowNumber)
  {
    const headerRowCount = this.getHeaderRowCount();
    if (rowNumber <= headerRowCount) return null;

    const lastCol = this.sheet.getLastColumn();
    if (lastCol === 0) throw new Error(`Sheet "${this.sheetName}" has no data.`);

    const range = this.sheet.getRange(rowNumber, 1, 1, lastCol);
    return range.getValues()[0];
  }

  /**
   * Returns a row of data as an object mapping column names to values.
   *
   * @param {number} rowNumber - The row number (1-based).
   * @returns {Object<string, any> || null} Object of column name → cell value, or null if within header.
   */
  getRowData(rowNumber)
  {
    const headerRowCount = this.getHeaderRowCount();
    if (rowNumber <= headerRowCount) return null;

    const colNumbers = this.getColumnNumbers();
    const result = {};

    for (const colName in colNumbers) {
      const col = colNumbers[colName];
      result[colName] = this.sheet.getRange(rowNumber, col).getValue();
    }

    return result;
  }

  /**
   * Retrieves a named range from the spreadsheet by its global name.
   *
   * @param {string} rangeName - The name of the named range.
   * @returns {Range|null} `Range` object, or null if not found.
   * Note: Returning null is the behavior of SpreadsheetApp.getRangeByName(), 
   * not of this method itself.
   */
  getNamedRange(rangeName)
  {
    return this.sheet.getParent().getRangeByName(rangeName);
  }

  /**
   * Returns the validation rule for a column name, if any.
   *
   * @param {string} colName - The logical column name.
   * @returns {GoogleAppsScript.Spreadsheet.DataValidation | null}
   */
  getColumnValidationRule(colName)
  {
    const ruleFn = SHEETCONFIG[this.sheetName]?.validationRules?.column;
    return typeof ruleFn === 'function' ? ruleFn(colName) : null;
  }

  /**
   * Returns the validation rule for a configured named range, if any.
   *
   * @param {string} rangeName - The logical range name.
   * @returns {GoogleAppsScript.Spreadsheet.DataValidation | null}
   */
  getRangeValidationRule(rangeName)
  {
    const ruleFn = SHEETCONFIG[this.sheetName]?.validationRules?.range;
    return typeof ruleFn === 'function' ? ruleFn(rangeName) : null;
  }

  /**
   * Retrieves the conditional formatting rule builder function for a given column name.
   *
   * @param {string} colName - The name of the column (e.g., "department").
   * @returns {(function(GoogleAppsScript.Spreadsheet.Sheet, SmartSheet): GoogleAppsScript.Spreadsheet.ConditionalFormatRuleBuilder[])|null}
   *   A function that returns conditional formatting rule builders, or null if not defined.
   */
  getColumnFormattingRule(colName)
  {
    const ruleFn = SHEETCONFIG[this.sheetName]?.formattingRules?.column;
    return typeof ruleFn === 'function' ? ruleFn(colName) : null;
  }

  /**
   * Retrieves the conditional formatting rule builder function for a named range.
   *
   * @param {string} rangeName - The name of the defined range (e.g., "selectDepartment").
   * @returns {(function(GoogleAppsScript.Spreadsheet.Sheet, SmartSheet): GoogleAppsScript.Spreadsheet.ConditionalFormatRuleBuilder[])|null}
   *   A function that returns conditional formatting rule builders, or null if not defined.
   */
  getRangeFormattingRule(rangeName)
  {
    const ruleFn = SHEETCONFIG[this.sheetName]?.formattingRules?.range;
    return typeof ruleFn === 'function' ? ruleFn(rangeName) : null;
  }

}
