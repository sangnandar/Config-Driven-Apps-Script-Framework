/*******************************************************
 **        GLOBAL VARIABLES AND CONFIGURATIONS        **
 *******************************************************/

/*
 * Rename this file to config.gs
 * and replace the placeholder values
 * with your actual values.
 */

const DEBUG = true; // set to false for production

var UI; // return null if called from script editor
try {
  UI = SpreadsheetApp.getUi();
} catch (e) {
  Logger.log('You are using script editor.');
}
const SS = SpreadsheetApp.getActiveSpreadsheet();

// === START: Lists ===

const departmentList = readOnlyObject({
  'HR': {
    bgColor: '#3FA7D6'
  },
  'Engineering': {
    bgColor: '#E65F5C'
  },
  'Sales': {
    bgColor: '#8DD694'
  },
  'Marketing': {
    bgColor: '#F2C94C'
  },
  'Finance': {
    bgColor: '#A66DD4'
  }
});

/**
 * Builds conditional formatting rule builders for `department`.
 *
 * @param {Object<string, {bgColor: string}>} valueList
 * @returns {GoogleAppsScript.Spreadsheet.ConditionalFormatRuleBuilder[]}
 */
const formatDepartment = (valueList) => {
  return Object.entries(valueList).map(([value, config]) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(value)
      .setBackground(config.bgColor)
  );
}

// === END: Lists ===


// === START: Configuration for Spreadsheets ===

const editors = [
  {
    email: '<AN_EMAIL_ADDRESS>',
    protector: true
  }
];

// === END: Configuration for Spreadsheets ===

// === START: Configuration for Sheets ===

// Sheet: 'Employees'
const SHEETNAME_EMPLOYEES = DEBUG
  ? 'Employees_dev' // for development & debugging
  : 'Employees'; // for production

// Sheet: <add more sheets...>

const SHEETCONFIG = readOnlyObject({

  [SHEETNAME_EMPLOYEES]: {
    layout: {
      headerRows: 4,
      columns: {
        A   : { name: 'name',       type: 'string' },
        B   : { name: 'age',        type: 'number' },
        C   : { name: 'joinDate',   type: 'date'   },
        D   : { name: 'department', type: 'string' },
        E   : { name: 'score',      type: 'number' },
        F   : {
          name: 'relativeScore',
          formula: 'IF(ISBLANK($score), "", $score / AVERAGE(E5:E) )',
          lock: true
        }
        // <add more columns...>
      },
      namedRanges: {
        'B1' : { name: 'selectDepartment', type: 'string' },
        'B2' : { name: 'selectScore',      type: 'string' },
        'D1' : {
          name: 'cellD1',
          formula: '$selectDepartment',
          lock: true
        },
        'E1' : {
          name: 'cellE1', // name can't be the same with custom-function-name or A1 notation
          args: ['selectDepartment'],
          formula: forE1, // use custom-function
          lock: true
        }
        // <add more namedRanges...>
      }
    },
    /** @type {Array<EditRule>} */
    onEditRules: [
      {
        condition: (e, smartSheet) => {
          return e.range.getA1Notation() === smartSheet.getNamedRange('selectDepartment').getA1Notation();
        },
        handler: selectDepartmentChange
      }
      // <add more rules...>
    ],
    validationRules: {
      column: (columnName) => {
        return {
          department: SpreadsheetApp.newDataValidation()
            .requireValueInList(Object.keys(departmentList), true)
            .setAllowInvalid(false)
            .build()

          // <add more rules for column...>
        }[columnName] || null;
      },
      range: (rangeName) => {
        return {
          selectDepartment: SpreadsheetApp.newDataValidation()
            .requireValueInList(Object.keys(departmentList), true)
            .setAllowInvalid(false)
            .build(),

          selectScore: SpreadsheetApp.newDataValidation()
            .requireValueInList(['0-20', '20-40', '40-60', '60-80', '80-100'], true)
            .setAllowInvalid(false)
            .build()

          // <add more rules for range...>
        }[rangeName] || null;
      }
    },
    formattingRules: {
      column: (columnName) => {
        return {
          department: () => formatDepartment(departmentList)

          // <add more rules for column...>
        }[columnName] || null;
      },
      range: (rangeName) => {
        return {
          selectDepartment: () => formatDepartment(departmentList)

          // <add more rules for range...>
        }[rangeName] || null;
      }
    },
    namedFunctions: {
      // not yet possible using Apps Script
    }

  }
  // <add more sheets...>
});

// === END: Configuration for Sheets ===


/**
 * A generic rule for handling various Google Apps Script events.
 *
 * @template T
 * @typedef {Object} Rule
 * @property {(e: T, smartSheet?: SmartSheet) => boolean} condition
 * @property {(e: T, smartSheet?: SmartSheet) => void} handler
 */

/**
 * Specific rule typedefs.
 * 
 * @typedef {Rule<GoogleAppsScript.Events.SheetsOnEdit>} EditRule
 * @typedef {Rule<GoogleAppsScript.Events.SheetsOnChange>} ChangeRule
 * @typedef {Rule<GoogleAppsScript.Events.SheetsOnOpen>} OpenRule
 */


