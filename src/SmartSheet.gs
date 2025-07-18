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
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object
   * @param {Object} configObject - (Optional) The configuration object (SHEETCONFIG or EXTERNAL_SHEETCONFIG)
   */
  constructor(sheet, configObject)
  {
    this.sheet = sheet;
    this.sheetName = sheet.getName();
    this.config = configObject || SHEETCONFIG;
    this.sheetConfig = this.config[this.sheetName] || {};

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
    return this.sheetConfig.layout?.headerRows ?? 0;
  }

  /**
   * Retrieves the column configuration for the current sheet from the global SHEETCONFIG.
   * 
   * @returns {Object<string, { name: string, type: string }>}
   *          An object mapping column letters to their configuration.
   */
  getColumnConfig()
  {
    return this.sheetConfig.layout?.columns ?? {};
  }

  /**
   * Returns the named range configuration for the current sheet from the global SHEETCONFIG.
   *
   * @returns {Object<string, { name: string, type: string }>}
   *          An object mapping range notations to their configuration.
   */
  getNamedRangeConfig()
  {
    return this.sheetConfig.layout?.namedRanges ?? {};
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

    const columnConfig = this.getColumnConfig();
    this._calculatedColumns = {};
    for (const columnLetter in columnConfig) {
      if (columnConfig[columnLetter].hasOwnProperty('formula')) {
        this._calculatedColumns[columnLetter] = columnConfig[columnLetter];
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

    const namedRangeConfig = this.getNamedRangeConfig();
    this._calculatedNamedRanges = {};
    for (const notation in namedRangeConfig) {
      if (namedRangeConfig[notation].hasOwnProperty('formula')) {
        this._calculatedNamedRanges[notation] = namedRangeConfig[notation];
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

    const columnConfig = this.getColumnConfig();
    this._columnLetters = {};
    for (const columnLetter in columnConfig) {
      const { name } = columnConfig[columnLetter];
      this._columnLetters[name] = columnLetter.toUpperCase();
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

    const columnConfig = this.getColumnConfig();
    this._columnTypes = {};
    for (const columnLetter in columnConfig) {
      const { name, type } = columnConfig[columnLetter];
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

    const columnLetters = this.getColumnLetters();
    this._columnNumbers = {};
    for (const columnName in columnLetters) {
      const columnLetter = columnLetters[columnName];
      this._columnNumbers[columnName] = columnLetter
        .split('')
        .reduce((total, char) => total * 26 + (char.charCodeAt(0) - 64), 0);
    }
    return this._columnNumbers;
  }

  /**
   * Returns a mapping of named range names to their A1Notations (in uppercase).
   *
   * @returns {Object.<string, string>} A map of namedRange → A1Notation.
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
   * @param {string} columnName - The name of the column as defined in config.
   * @returns {any[]} 1-D array of values from that column.
   */
  getColumnValues(columnName)
  {
    const headerRows = this.getHeaderRowCount();
    const columnNumber = this.getColumnNumbers()[columnName];

    if (!columnNumber) throw new Error(`Unknown column name: ${columnName}`);

    const maxRows = this.sheet.getMaxRows();
    const lastRow = this.sheet.getRange(maxRows, columnNumber)
                              .getNextDataCell(SpreadsheetApp.Direction.UP)
                              .getRow();

    // If there's no data below headerRows, return empty array
    if (lastRow <= headerRows) return [];

    const range = this.sheet.getRange(headerRows + 1, columnNumber, lastRow - headerRows, 1);
    return range.getValues().flat();
  }

  /**
   * Sets values in a column by column name, starting after the header rows.
   * Automatically expands the sheet if needed.
   *
   * @param {string} columnName - The logical column name.
   * @param {any[][]} value - A 2D array of values to write into the column.
   * @returns {void}
   * @throws {Error} If column name is invalid or value is not a 2D array.
   */
  setColumnValues(columnName, value)
  {
    if (!Array.isArray(value) || !Array.isArray(value[0])) {
      throw new Error('Value must be a 2D array.');
    }

    const headerRowCount = this.getHeaderRowCount();
    const columnLetters = this.getColumnLetters();
    const columnLetter = columnLetters[columnName];

    if (!columnLetter) {
      throw new Error(`Unknown column name: ${columnName}`);
    }

    const startRow = headerRowCount + 1;
    const endRow = headerRowCount + value.length;
    const maxRow = this.sheet.getMaxRows();

    if (endRow > maxRow) {
      this.sheet.insertRowsAfter(maxRow, endRow - maxRow);
    }

    const rangeNotation = `${columnLetter}${startRow}:${columnLetter}${endRow}`;
    this.sheet.getRange(rangeNotation).setValues(value);
  }

  /**
   * Returns all values in the specified row as a 1D array.
   * Returns null if the row is within the configured header rows or if the sheet has no data.
   * Throws an error if the sheet has no data.
   * 
   * @param {number} rowNumber - The 1-based row-number from the sheet.
   * @returns {any[] | null} 1D array of values from the row, or null if within header.
   */
  getRowValues(rowNumber)
  {
    const headerRowCount = this.getHeaderRowCount();
    const lastColumn = this.sheet.getLastColumn();
    if (
      rowNumber <= headerRowCount ||
      lastColumn === 0 // sheet contains no data, not even header
    ) return null;

    return this.sheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0];
  }

  /**
   * Returns a row of data as an object mapping column names to values.
   *
   * @param {number} rowNumber - The 1-based row-number from the sheet.
   * @returns {Object<string, any> | null} Object of columnName → value, or null if within header.
   */
  getRowData(rowNumber)
  {
    const headerRowCount = this.getHeaderRowCount();
    if (rowNumber <= headerRowCount) return null;

    const columnNumbers = this.getColumnNumbers();
    const rowValues = this.getRowValues(rowNumber);
    if (!rowValues) return null; // just chain the null

    const result = {};
    for (const columnName in columnNumbers) {
      const columnNumber = columnNumbers[columnName];
      result[columnName] = rowValues[columnNumber - 1];
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
   * @param {string} columnName - The logical column name.
   * @returns {GoogleAppsScript.Spreadsheet.DataValidation | null}
   */
  getColumnValidationRule(columnName)
  {
    const ruleFn = this.sheetConfig.validationRules?.column;
    return typeof ruleFn === 'function' ? ruleFn(columnName) : null;
  }

  /**
   * Returns the validation rule for a configured named range, if any.
   *
   * @param {string} rangeName - The logical range name.
   * @returns {GoogleAppsScript.Spreadsheet.DataValidation | null}
   */
  getRangeValidationRule(rangeName)
  {
    const ruleFn = this.sheetConfig.validationRules?.range;
    return typeof ruleFn === 'function' ? ruleFn(rangeName) : null;
  }

  /**
   * Retrieves the conditional formatting rule builder function for a given column name.
   *
   * @param {string} columnName - The name of the column (e.g., "department").
   * @returns {(function(GoogleAppsScript.Spreadsheet.Sheet, SmartSheet): GoogleAppsScript.Spreadsheet.ConditionalFormatRuleBuilder[])|null}
   *   A function that returns conditional formatting rule builders, or null if not defined.
   */
  getColumnFormattingRule(columnName)
  {
    const ruleFn = this.sheetConfig.formattingRules?.column;
    return typeof ruleFn === 'function' ? ruleFn(columnName) : null;
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
    const ruleFn = this.sheetConfig.formattingRules?.range;
    return typeof ruleFn === 'function' ? ruleFn(rangeName) : null;
  }

}
