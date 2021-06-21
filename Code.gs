const recipients = 'dd@example.com'; // comma-delimited list of recipients

/**
 * Admin functions
 */

/**
 * Reset properties to empty, to ignore changes
 */
 function resetSheetUpdate() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('updates', '{}');
}

/**
 * Save a modification flag to script properties
 * @param {int} sheetnum number of sheet cell belongts to
 * @param {string} activeCell e.g. 'A1', 'C3'
 */
 function saveSheetUpdate(sheetnum, activeCell) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentUpdates = JSON.parse(scriptProperties.getProperty('updates'));
  const updates = (currentUpdates !== null)?currentUpdates:{};
  if (updates[sheetnum] !== undefined) {
    updates[sheetnum] = [...new Set([ ... updates[sheetnum], activeCell ])].sort();
  } else {
    updates[sheetnum] = [ activeCell ];
  }

  try {
    scriptProperties.setProperty('updates', JSON.stringify(updates));
  } catch (err) {
    Logger.log(err.message);
  }
}

/**
 * Helper functions
 */

/**
 * Returns numeric equivalent column , e.g. B = 2, J = 10
 * @param {string} colA1 the alphabetic column name, e.g. 'A', 'J', 'AB'
 * @returns {int} column equivalent 1, 10, 28
 */
 function colA1ToIndex(colA1) {
  const A = "A".charCodeAt(0);
  const radix = "Z".charCodeAt(0) - A + 1;
  const l = colA1.length;
  let sum = 0;
  if (typeof colA1 !== 'string' || !/^[A-Z]+$/.test(colA1)) {
    throw new Error('Expected column label');
  }
  for (i = 0 ; i < l ; i++) {
    chr = colA1.charCodeAt(i);
    sum = sum * radix + chr - A + 1
  }
  return sum;
}

/**
 * Checks for valid row number
 * @param {int} rowA1 e.g. 1, 19, 34
 * @returns {int} e.g. 1, 19, 34
 */
function rowA1ToIndex(rowA1) {
  const index = parseInt(rowA1, 10);
  if(isNaN(index)) {
    throw new Error('Expected row number');
  }
  return index;
}

/**
 * Convert A1 reference to a row, column tuple
 * @param {string} cellA1 e.g. A19
 * @returns {object} tuple e.g. {'row':19,'col':1}
 */
function cellA1ToIndex(cellA1) {
  const match = cellA1.match(/^\$?([A-Z]+)\$?(\d+)$/);
  if(!match) {
    throw new Error('Invalid cell reference');
  }
  return {
    row: rowA1ToIndex(match[2]),
    col: colA1ToIndex(match[1])
  };
}

/**
 * Convert numeric coordinates to A1 notation
 * @param {int} row e.g. 1, 12, 42
 * @param {int} column e.g. 2, 13, 19
 * @returns {string} A1 reference e.g. 'B1', 'M12', 'R42'
 */
function R1C1toA1 (row, column) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const base = chars.length;
  let columnRef = '';

  if (column < 1) {
    columnRef = chars[0];
  } else {
    let maxRoot = 0;
    while (base**(maxRoot + 1) < column) {
      maxRoot++;
    }

    let remainder = column;
    for (let root = maxRoot; root >= 0; root--) {
      const value = Math.floor(remainder / base**root);
      remainder -= (value * base**root);
      columnRef += chars[value - 1];
    }
  }

  // Use Math.max to ensure minimum row is 1
  return `${columnRef}${Math.max(row, 1)}`;
};

/**
 * Calculate list of spanned rows, and list of cells within range
 * @param {string} rangeSpecifier, e.g. A2:C3
 * @returns {object} { cells:['A2','A3','B2','B3','C2','C3' ] , rows:[1,2,3] }
 */
function normaliseRange(rangeSpecifier) {
  const cells = [];
  const rows = [];

  if(rangeSpecifier.indexOf(':') > -1) {
    const [from, until] = rangeSpecifier.split(':');
    const {row:startRow , col:startCol} = cellA1ToIndex(from);
    const {row:endRow , col:endCol} = cellA1ToIndex(until);

    for(i = startRow ; i <= endRow ; i++) {
      rows.push(i); // push row
      for(j = startCol ; j <= endCol ; j++) {
        cells.push(R1C1toA1(i,j));
      }
    }
  } else {
    // we have a single cell.
    rows.push(Number(rangeSpecifier.replace(/[^0-9]/gi,'')));
    cells.push(rangeSpecifier);
  }
  return { cells , rows };
}

/**
 * Convert list of rows into list of contiguous ranges.
 * @param {array} rows e.g. [ 1, 2, 3, 4, 8 ]
 * @returns {array} [ [ 1, 4 ], [ 8, 8 ] ]
 */
function convertToRanges (rows) {
  // ranges will be an array of arrays, each inner array will
  // have 2 dimensions, representing the start/end of a range
  // we want to initialize our first range to rows[0], rows[0]
  const ranges = [[rows[0], rows[0]]]
  // last index we accessed (so we know which range to update)
  let lastIndex = 0;

  for(i = 1; i < rows.length; i++) {
    // if the current element is 1 away from the end of whichever range
    // we're currently in
    if (rows[i] - ranges[lastIndex][1] === 1) {
      // update the end of that range to be this number
      ranges[lastIndex][1] = rows[i];
    } else {
      // otherwise, add a new range to ranges
      ranges[++lastIndex] = [rows[i], rows[i]];
    }
  }
  return ranges;
}

/**
 * Process modifications properties
 */
function processUpdates() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const updates = JSON.parse(scriptProperties.getProperty('updates'));

  // only process if we have updates to send!
  if(Object.keys(updates).length > 0) {
    // figure out which rows to show, and which cells to highlight
    const cells = {};
    const rows = {};
    let html = '<p>Link to this sheet: '+SpreadsheetApp.getActiveSpreadsheet().getUrl()+'</p>';

    // first, process ALL the updates for all sheets
    Object.keys(updates).forEach((sheetNum, i) => {
      updates[sheetNum].forEach((update) => {
        // we only need to normalise ranges, not rows or columns
        const newItems = normaliseRange(update);
        cells[sheetNum] = [...new Set([ ... (cells[sheetNum] !== undefined ? cells[sheetNum] : []), ...newItems.cells ])].sort();
        rows[sheetNum] = [...new Set([ ... (rows[sheetNum] !== undefined ? rows[sheetNum] : []), ...newItems.rows ])].sort();
      });
    });

    // then summaries the updates
    Object.keys(updates).forEach((sheetNum, i) => {
      const lastColumn = cells[sheetNum].map(e => { return e.replace(/[0-9]/gi,''); }).sort().reverse()[0];
      const ranges = convertToRanges(rows[sheetNum]).map(e => { return 'A'+e[0]+':'+lastColumn+e[1]});
      html += '<h2>Updated: '+SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheetNum].getSheetName()+'</h2>';
      ranges.forEach((range) => {
        html += '<h3>'+range+'</h3>';
        html += getHtmlTable(Number(sheetNum),range,cells[sheetNum]);
      })
    });

    // send the update email, and then wipe the slate...
    sendEmailAlert(html);
    resetSheetUpdate();
  }
}

/**
 *
 */
function receiveUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeCell = ss.getActiveSheet().getActiveCell().getA1Notation();
  const sheetnum = ss.getActiveSheet().getIndex() - 1;
  saveSheetUpdate(sheetnum, activeCell);
}

/**
 * Send an html email to recipients list
 * @param {string} html layout code to include in email
 */
function sendEmailAlert(html) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetname = ss.getName();
  const subject = 'Sheet Update - ' + sheetname;
  const body = 'Your sheet, ' + sheetname + ', has been. Check file - ' + ss.getUrl();
  MailApp.sendEmail(recipients,subject, body, {htmlBody : html});
};

/**
 * Return a string containing an HTML table representation
 * of the given range, preserving style settings.
 * @param {int} sheetNumber sheet to use e.g. 0
 * @param {string} rangeSpecifier range to create table from e.g. 'A1:F4'
 * @param {array} cellsToHighlight cells to highlight e.g. [ 'A1', 'A2', 'D4', 'E4', 'F8' ]
 * @returns {string} html code for the table
 */
function getHtmlTable(sheetNumber,rangeSpecifier, cellsToHighlight){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheetNumber];
  const range = sheet.getRange(rangeSpecifier);
  const ss = sheet.getParent();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const lastRow = range.getLastRow();
  const lastCol = range.getLastColumn();

  // Read table contents
  const numberFormats = range.getNumberFormats();
  const data = range.getValues();

  // Get css style attributes from range
  const fontColors = range.getFontColors();
  const backgrounds = range.getBackgrounds();
  const fontFamilies = range.getFontFamilies();
  const fontSizes = range.getFontSizes();
  const fontLines = range.getFontLines();
  const fontWeights = range.getFontWeights();
  const horizontalAlignments = range.getHorizontalAlignments();
  const verticalAlignments = range.getVerticalAlignments();

  // Get column widths in pixels
  const colWidths = [];
  for(col = startCol; col <= lastCol; col++) {
    colWidths.push(sheet.getColumnWidth(col));
  }
  // Get Row heights in pixels
  const rowHeights = [];
  for(row = startRow; row <= lastRow; row++) {
    rowHeights.push(sheet.getRowHeight(row));
  }

  // Future consideration...
  // const numberFormats = schedRange.getNumberFormats();

  // Build HTML Table, with inline styling for each cell
  const tableFormat = 'style="border:1.5px solid black;border-collapse:collapse;text-align:center" border = 1.5 cellpadding = 5';
  const html = ['<table '+tableFormat+'>'];
  // Column widths appear outside of table rows
  for (col=0;col<colWidths.length;col++) {
    html.push('<col width="'+colWidths[col]+'">')
  }
  // Populate rows
  for (row=0;row<data.length;row++) {
    html.push('<tr height="'+rowHeights[row]+'">');
    for (col=0;col<data[row].length;col++) {
      // Get formatted data
      let cellText = data[row][col];
      if (cellText instanceof Date) {
        cellText = Utilities.formatDate(
                     cellText,
                     ss.getSpreadsheetTimeZone(),
                     'MMM/d EEE');
      }
      const style = 'style="'
                + 'color: ' + fontColors[row][col]+'; '
                + 'font-family: ' + fontFamilies[row][col]+'; '
                + 'font-size: ' + fontSizes[row][col]+'; '
                + 'font-weight: ' + fontWeights[row][col]+'; '
                + 'background-color: ' + backgrounds[row][col]+'; '
                + 'text-align: ' + horizontalAlignments[row][col]+'; '
                + 'vertical-align: ' + verticalAlignments[row][col]+'; '
                + (cellsToHighlight.indexOf(R1C1toA1(row+startRow,col+startCol)) > -1 ? 'background-color: lightgreen' : '')
                +'"';
      html.push('<td ' + style + '>'
                +cellText
                +'</td>');
    }
    html.push('</tr>');
  }
  html.push('</table>');

  return html.join('');
}

/**
 * Debug functions
 */

/**
 * Create a set of dummy values to test
 */
 function createTestValues() {
  saveSheetUpdate(0, 'A1:B4');
  saveSheetUpdate(1, 'A4:C4');
  saveSheetUpdate(1, 'A2:C5');
  saveSheetUpdate(0, 'E4');
  saveSheetUpdate(0, 'D4');
  saveSheetUpdate(0, 'F8');
}

/**
 * List updates to console
 */
function listUpdates() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentUpdates = JSON.parse(scriptProperties.getProperty('updates'));
  console.log(currentUpdates);
}
