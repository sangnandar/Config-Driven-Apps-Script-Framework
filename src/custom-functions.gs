/*******************************************************
 **                 CUSTOM FUNCTIONS                  **
 *******************************************************/

/**
 * Array-returning function to be written into Sheets.
 *
 * @param {string} arg - A string representing a cell or range in A1 notation (e.g., "A1", "B2:C3").
 * @returns {string[][]} A 2D array.
 */
function forE1(arg)
{
  return [
    [
     ` ${arg}-1`,
     ` ${arg}-2`,
     ` ${arg}-3`
    ]
  ];
}