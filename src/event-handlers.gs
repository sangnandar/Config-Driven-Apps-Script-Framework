/*******************************************************
 **               TRIGGER/EVENT HANDLERS              **
 *******************************************************/

/**
 * Handler for `selectDepartment` edit.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - Event object.
 * @param {SmartSheet} smartSheet - The SmartSheet instance for the edited sheet.
 * @returns {void}
 */
function selectDepartmentChange(e, smartSheet)
{
  const sheet = e.range.getSheet();
  sheet.getRange('C1').setValue(e.value);
}