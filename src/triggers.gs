/*******************************************************
 **      SIMPLE, INSTALLABLE, AND HTTP TRIGGERS       **
 *******************************************************/

/**
 * Run column-based and range-based validations.
 *
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e - Event object.
 * @returns {void}
 */
function onOpen(e)
{
  const ss = e.source;

  if (DEBUG) clearAll(ss); // reset the spreadsheet

  const currentEditors = new Set(ss.getEditors().map(e => e.getEmail()));
  const newEditors = editors
    .map(e => e.email)
    .filter(email => !currentEditors.has(email));

  if (newEditors.length > 0) {
    try {
      ss.addEditors(newEditors);
    } catch (e) {
      showAlert(`Failed to add editors: ${e.message}`); // emails might not have associated google accounts
    }
  }

  // validate SHEETCONFIG first
  const valid = validateSHEETCONFIG();
  if (!valid) {
    showAlert('Not valid. Check the log.');
    return;
  }

  for (const sheetName in SHEETCONFIG) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const smartSheet = new SmartSheet(sheet);

    setNamedRanges(smartSheet); // run this first

    applyCalculatedColumns(smartSheet);
    applyCalculatedNamedRanges(smartSheet)
    applyValidationRules(smartSheet);
    applyFormattingRules(smartSheet);
  }
}

/**
 * Aggregator for all edit events.
 * Calls the appropriate handler based on event conditions.
 * 
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - event object.
 * @returns {void}
 */
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
        return; // or continue for fallthrough support (run multiple handlers that share the same condition)
      }
    }
  }

  return; // no match
}
