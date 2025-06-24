# Config-Based Apps Script Framework for Google Sheets

A modular, config-driven framework that simplifies building scalable and maintainable Google Sheets solutions using Google Apps Script. By centralizing logic in a structured `SHEETCONFIG` object and utilizing a powerful `SmartSheet` class, this framework helps you manage layout, data rules, and automation logic consistently across multiple sheets.

This project is the final iteration of concepts developed in previous works:

- [📊 Dynamic Google Sheets Layout](https://github.com/sangnandar/Dynamic-Google-Sheets-Layout)
- [📋 Lookup Table Pattern in Apps Script](https://github.com/sangnandar/Lookup-Table-Pattern)
- [⚡ Event Dispatcher in Apps Script](https://github.com/sangnandar/Event-Dispatcher-in-Apps-Script)


## ✨ Features

- 🧩 **Centralized Configuration (`SHEETCONFIG`)**
  - Define columns, named ranges, validation rules, formatting rules, and edit behavior in one place.
- 🧠 **SmartSheet Class**
  - A utility wrapper that interprets `SHEETCONFIG` and simplifies reading/writing data.
- 🛡️ **Validation & Conditional Formatting**
  - Attach rules by column or named range with support for list values and styling.
- ⚡ **Calculated Columns & Named Ranges**
  - Use inline formulas or functions with `$var` injection to populate dynamic values.
- 🔐 **Protected Fields**
  - Mark columns or named ranges as locked to prevent user editing.
- ✅ **Rule-Based onEdit Handlers**
  - Handle edit events declaratively using rule conditions and handlers.
- 🔧 **Debug Mode Support**
  - Switch between production and development sheets easily.


## 📁 Project Structure

```

/src
├── config.gs             # Global SHEETCONFIG and constants
├── SmartSheet.gs         # SmartSheet utility class
├── utils.gs              # Utility and helper functions
├── triggers.gs           # Entry points for onOpen, onEdit, installable, etc.
├── custom-functions.gs   # Custom spreadsheet functions
└── event-handlers.gs     # Named event handlers used in config

````


## 🛠️ Getting Started

1. **Copy the contents of `/src` into your Apps Script project.**
2. Define your sheet layout and logic in `config.gs` using the `SHEETCONFIG` object.
3. Customize `triggers.gs` to apply your configuration and handle edits.
4. Add formulas, validations, and formatters as needed.

---

## 📦 Example: Employees Sheet

```js
const SHEETCONFIG = readOnlyObject({
  Employees: {
    layout: {
      headerRows: 4,
      columns: {
        A: { name: 'name',       type: 'string' },
        B: { name: 'age',        type: 'number' },
        C: { name: 'joinDate',   type: 'date' },
        D: { name: 'department', type: 'string' },
        E: { name: 'score',      type: 'number' },
        F: { // calculated column
          name: 'relativeScore',
          formula: 'IF(ISBLANK($score), "", $score / AVERAGE(E5:E))',
          lock: true
        }
      },
      namedRanges: {
        'B1': { name: 'selectDepartment', type: 'string' },
        'E1': { // calculated named-range
          name: 'cellE1',
          args: ['selectDepartment'],
          formula: forE1,
          lock: true
        }
      }
    },
    onEditRules: [
      {
        condition: (e, smartSheet) => {
          return e.range.getA1Notation() === smartSheet.getNamedRange('selectDepartment').getA1Notation();
        },
        handler: selectDepartmentChange
      }
    ],
    validationRules: {
      column: (col) => ({
        department: SpreadsheetApp.newDataValidation()
          .requireValueInList(['HR', 'Engineering', 'Sales'], true)
          .setAllowInvalid(false)
          .build()
      }[col] || null),
      range: (range) => ({
        selectDepartment: SpreadsheetApp.newDataValidation()
          .requireValueInList(['HR', 'Engineering', 'Sales'], true)
          .setAllowInvalid(false)
          .build()
      }[range] || null)
    },
    formattingRules: {
      column: (col) => ({
        department: () => formatDepartment(departmentList)
      }[col] || null)
    }
  }
});
````


## 🚀 Usage in Trigger Functions

The framework uses standard Apps Script triggers defined in `triggers.gs`:

```js
function onOpen(e)
{
  const ss = e.source;

  if (DEBUG) clearAll(ss); // optional: reset the spreadsheet in dev

  const currentEditors = new Set(ss.getEditors().map(e => e.getEmail()));
  const newEditors = editors
    .map(e => e.email)
    .filter(email => !currentEditors.has(email));

  if (newEditors.length > 0) {
    try {
      ss.addEditors(newEditors);
    } catch (e) {
      showAlert(`Failed to add editors: ${e.message}`);
    }
  }

  if (!validateSHEETCONFIG()) {
    showAlert('Not valid. Check the log.');
    return;
  }

  for (const sheetName in SHEETCONFIG) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const smartSheet = new SmartSheet(sheet);
    setNamedRanges(smartSheet);
    applyCalculatedColumns(smartSheet);
    applyCalculatedNamedRanges(smartSheet);
    applyValidationRules(smartSheet);
    applyFormattingRules(smartSheet);
  }
}

function onEdit(e)
{
  const sheet = e.range.getSheet();
  const sheetName = sheet.getSheetName();

  const rules = SHEETCONFIG[sheetName]?.onEditRules;
  if (rules) {
    const smartSheet = new SmartSheet(sheet);
    for (const rule of rules) {
      if (rule.condition(e, smartSheet)) {
        rule.handler(e, smartSheet);
        return;
      }
    }
  }
}
```

These triggers:

* Initialize the sheet when opened
* Handle dynamic logic on user edits based on your config


## 📚 Documentation

* **`SHEETCONFIG`**

  * Keys:

    * `layout` → define structure: `columns`, `namedRanges`, `headerRows`
    * `onEditRules` → define edit logic conditionally
    * `validationRules` → dynamic input validation
    * `formattingRules` → conditional formatting to be applied
    * `namedFunctions` → for future support

* **SmartSheet Methods**

  * `getRowData(rowNumber)`
  * `getColumnValues(colName)`
  * `getNamedRange(name)`
  * `getCalculatedColumns()`
  * `getColumnValidationRule(name)`
  * `getColumnFormattingRule(name)`
    *(and many more)*


## 🧠 Design Philosophy

* **Declarative > Imperative**: Let config describe what your sheet is and how it behaves.
* **Reusable**: Avoid boilerplate logic across sheets.
* **Maintainable**: Add/edit/remove sheet logic in config, not in multiple functions.
* **Transparent**: Anyone reading the config knows the sheet structure and logic.


## 🙌 Credits

Framework by [Sunandar Gusti](https://github.com/sangnandar), based on experience building complex automation for Google Sheets using clean, modular patterns.


## 🔭 Planned Improvements

This project is actively maintained. Here are a few next development goals:

- ✅ **Improve `validateSHEETCONFIG`**
  - Add deeper integrity checks (e.g. missing column names, missing formula for calculated columns, missing handler functions).
  
- ✅ **Implement Config Caching**
  - Use `CacheService` to store a serialized version of the `SHEETCONFIG` object to improve performance, especially in large spreadsheet.
