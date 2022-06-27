/*
A set of classes to convert Google Sheets data from a given sheet into HTML so that cell formatting is preserved as much as possible.
The code preserves text properties such as colour, bold and italics as well as cell background and merged cells (columns and/or rows).
It turns cell notes into text that appears with the onhover action.

Currently it has no way to preserve sparklines.

 */


/**
 * Collect all the relevant cell attributes of a Google Sheets range into a data structure for use in other classes.
 * Loops over all the cells in the range and builds data structures of the attributes of each cell.
 * Merged ranges are treated differently from attributes that apply too single cells and aremapped to HTML rowspan and colspan attributes.
 */
class RangeCellAttributes {
  /**
   * Take the given range object and extract attributes from its constituent cells.
   * 
   * @param {Range} rng - A google Sheets Range object containing the cells for HTML conversion.
   */
  constructor(rng) {
    this.rng = rng;
    this.rngCellAttributes = this.createListofCellAttributes();
    this.mergedRngs = this.rng.getMergedRanges();
    this.mergedRngAttributes = this.createMapOfMergedRngAttributes();
  }

  /**
   * Get an instance variable containing cell attributes created by createListofCellAttributes().
   * Returns am array of arrays wth the inner array composed of objects where the objects are attribute-attribute value pairs.
   * Client code should need only this method to use the class.
   * 
   * @return {array<array>} 
   */
  getRngCellAttributes() {
    return this.rngCellAttributes;
  }

  /**
   * Get an instance variable containing  data structure created by
   * createMapOfMergedRngAttributes().
   * 
   * @todo Simpler to call createMapOfMergedRngAttributes() directly?
   * 
   * @return {object} - keys: first cell in merged range, values: merged range attributes object.
   */
  getMapOfMergedRngAttributes() {
    return this.mergedRngAttributes;
  }

  /**
   * Loop over all the cells of the range and build an an array of arrays where the inner arrays are composed of 
   * objects that map cell attribute names to cell attributes.
   * 
   * @return {array<array>}
   */
  createListofCellAttributes() {
    let rngCellAttributes = [];
    const rngRowCount = this.rng.getNumRows();
    const rngColCount = this.rng.getNumColumns();
    for (let i = 1; i <= rngRowCount; i++) {
      let rowAttrs = [];
      for (let j = 1; j <= rngColCount; j++) {
        let cell = this.rng.getCell(i, j);
        let cellAttrs = {};
        cellAttrs["address"] = cell.getA1Notation();
        cellAttrs["rangeRowIndex"] = i - 1;
        if (cellAttrs["rangeRowIndex"] === 0) {
          cellAttrs["htmlCellType"] = "th";
        } else {
          cellAttrs["htmlCellType"] = "td";
        }
        cellAttrs["rangeColIndex"] = j - 1;
        let cellType = typeof cell.getValue();
        cellAttrs["sheetCellType"] = cellType;
        cellAttrs["value"] = cell.getDisplayValue();
        cellAttrs["backgroundColor"] = cell.getBackground();
        cellAttrs["color"] = cell.getFontColor();
        cellAttrs["isMerged"] = cell.isPartOfMerge();
        cellAttrs["cellNote"] = cell.getNote();
        cellAttrs["fontStyle"] = cell.getFontStyle();
        cellAttrs["fontWeight"] = cell.getFontWeight();
        cellAttrs["textAlign"] = cell.getHorizontalAlignment();
        rowAttrs.push(cellAttrs);
      }
      rngCellAttributes.push(rowAttrs);
    }
    return rngCellAttributes;
  }

  /**
   * Create an array of merged cell addresses for an input range.
   * Used by populateMergedRngAttrs() to determine the merged range address.
   * 
   * @param {Range} mergedRng - An array of merged range addresses
   */
  getCellAddressesToSpan(mergedRng) {
    const rngRowCount = mergedRng.getNumRows();
    const rngColCount = mergedRng.getNumColumns();
    const firstCellAddress = mergedRng.getCell(1, 1).getA1Notation();
    let cellAddressesToSpan = [];
    for (let i = 1; i <= rngRowCount; i++) {
      for (let j = 1; j <= rngColCount; j++) {
        let cellAddress = mergedRng.getCell(i, j).getA1Notation();
        if (cellAddress !== firstCellAddress) {
          cellAddressesToSpan.push(cellAddress)
        }
      }
    }
    return cellAddressesToSpan;
  }

  /**
   * Build an object for a merged range that stores the information required to
   * convert the spreadsheet range merge notation into one that can be used in HTML
   * for colspan and rowspan.
   * 
   * @param {Range} mergedRng - A merged range object from which to extract required information
   * 
   * @return {object}
   */
  populateMergedRngAttrs(mergedRng) {
    let mergedRngAttrs = {}
    mergedRngAttrs["address"] = mergedRng.getA1Notation();
    mergedRngAttrs["rowSpanCount"] = mergedRng.getNumRows();
    mergedRngAttrs["colSpanCount"] = mergedRng.getNumColumns();
    mergedRngAttrs["firstCellAddress"] = mergedRng.getCell(1, 1).getA1Notation();
    mergedRngAttrs["cellsToSpan"] = this.getCellAddressesToSpan(mergedRng);
    return mergedRngAttrs;
  }

  /**
   * Loops over all merged ranges, calls populateMergedRngAttrs() for each merged range
   * and builds an object that uses the first cell of of the merged as a key that mas to marged range attributes.
   * 
   * @return {object} - keys: first cell in merged range, values: merged range attributes object.
   */ 
  createMapOfMergedRngAttributes() {
    let mergedRngAttrsAll = {};
    for (const mergedRng of this.mergedRngs) {
      const mergedRngAttrs = this.populateMergedRngAttrs(mergedRng);
      const firstCellAddress = mergedRngAttrs["firstCellAddress"];
      mergedRngAttrsAll[firstCellAddress] = mergedRngAttrs;
    }
    return mergedRngAttrsAll;
  }
}

/** Testing function for class  RangeCellAttributes
 * 
 * @return {void}
*/
function runRangeCellAttributes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HTML");
  const rng = sheet.getDataRange();
  const rngCellAttrs = new RangeCellAttributes(rng);
  console.log(rngCellAttrs.getRngCellAttributes()[2][5]);
  console.log(rngCellAttrs.getMapOfMergedRngAttributes());
}


/**
 * Transforms the data structures created by class RangeCellAttributes to include merged cell
 * information with other cell attributes and sets a primary object identifier for data structures
 * 
 * @todo Explore re-factoring and consider merging with class RangeCellAttributes.
 */
class AttrObjListForHtmlTableMaker {
  /**
   * Takes two data structures generated by RangeCellAttributes and two strings to set as the
   * primary identifiers for normal cell attributes and merged cell attributes.
   * 
   * @param {array} cellAttrsObjList
   * @param {object} mergedRngAttrsMap - keys: first cell in merged range, values: merged range attributes object.
   * @param {string} primaryCellAttrObjID - The key name used to map to cell attributes.
   * @param {string} primaryMergedRngAttrObjID - The key used to map to merged cell attributes
   */
  constructor(cellAttrsObjList,
    mergedRngAttrsMap,
    primaryCellAttrObjID,
    primaryMergedRngAttrObjID) {
    this.cellAttrsObjList = cellAttrsObjList;
    this.mergedRngAttrsMap = mergedRngAttrsMap;
    this.primaryCellAttrObjID = primaryCellAttrObjID
    this.primaryMergedRngAttrObjID = primaryMergedRngAttrObjID;
  }

  /**
   * Add merged info to array-of-arrays where there is a match on cell address with the merged info map.
   * Mutates this.cellAttrsObjList in place by doing a look-up to determine if a cell is part of a merged range.
   * If it is, it adds the merged information to the object in the inner array.
   * 
   * @todo The code is quite complex. Can this be simplified? 
   * 
   * @return {array} - Mutated version of the input.
   */
  addMergedInfoToCellAttrsObjList() {
    const rowCount = this.cellAttrsObjList.length;
    const colCount = this.cellAttrsObjList[0].length;
    for (let i = 0; i < rowCount; i++) {
      for (let j = 0; j < colCount; j++) {
        let cellAttrsObj = this.cellAttrsObjList[i][j];
        if (this.mergedRngAttrsMap.hasOwnProperty(cellAttrsObj[this.primaryCellAttrObjID])) {
          let mergedInfo = cellAttrsObj[this.primaryCellAttrObjID];
          cellAttrsObj["mergedInfo"] = this.mergedRngAttrsMap[mergedInfo];
          this.cellAttrsObjList[i][j] = cellAttrsObj
        }
      }
    }
    return this.cellAttrsObjList;
  }

}

/**
 * Testing function for class AttrObjListForHtmlTableMaker.
 * 
 * @return {void}
 */
function runAttrObjListForHtmlTableMaker() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HTML");
  const rng = sheet.getDataRange();
  const rngCellAttrs = new RangeCellAttributes(rng);
  const cellAttrsObjList = rngCellAttrs.getRngCellAttributes();
  const mergedRngAttrsMap = rngCellAttrs.getMapOfMergedRngAttributes();
  const primaryCellAttrObjID = "address";
  const primaryMergedRnfAttrObjID = "firstCellAddress";
  const attrObjListForHtml = new AttrObjListForHtmlTableMaker(cellAttrsObjList,
    mergedRngAttrsMap,
    primaryCellAttrObjID,
    primaryMergedRnfAttrObjID);

  const mergedInfoAddedToCellAttrsObjList = attrObjListForHtml.addMergedInfoToCellAttrsObjList();
  console.log(mergedInfoAddedToCellAttrsObjList[0]);
}

/**
 * Format a single cell entry in the data structure returned by AttrObjListForHtmlTableMaker.addMergedInfoToCellAttrsObjList().
 * as a HTML table cell.
 *
 */
class CellTransform {
  /**
   * Use the provided arguments to set instance variables and call methods to generate a formatted HTML table cell
   * 
   * @param {object} cellAttrs - An object from containing a range cell address key mapped to am object of cell attributes.
   * @param {object} lookupKeys - an object where the keys are attribute names and the values are whether they are required.
   */
  constructor(cellAttrs, lookupKeys) {
    this.cellAttrs = cellAttrs;
    this.lookupKeys = lookupKeys;
    this.tableCellType = this.getTableCellType();
    this.firstMergedCell = this.isFirstMergedCell();
    this.mergedToSpan = this.isMergedToSpan();
    this.formattedCell = this.formatCell();
  }

  /**
   * Check if the give attribute name is present in the lookupKeys object.
   * 
   * @param {string} attrName - The attribute name to check
   * @throw Error if attribute not found.
   * @return {string} 
   */
  getCellAttrValue(attrName) {
    if (this.lookupKeys.hasOwnProperty(attrName)) {
      return this.cellAttrs[attrName];
    }
    throw `No such attribute ${attrName}`;
  }

  /**
   *  Get the tpye of cell for the HTML table - in HTML format it is th or tr (heder or row cell).
   * 
   * @return {string}
   */
  getTableCellType() {
    const attrName = "htmlCellType";
    return this.getCellAttrValue(attrName);
  }

  /**
   * Determine if the cell is part of a merged range.
   * 
   * @return {boolean}
   */
  isFirstMergedCell() {
    const attrName = "mergedInfo";
    if (this.getCellAttrValue(attrName)) {
      return true;
    }
    return false;
  }

  /**
   * Determine if it is a cell that is parted of a merged range but not the first cell in this range.
   * 
   * @return {boolean}
   */ 
  isMergedToSpan() {
    const attrName = "isMerged";
    return (this.getCellAttrValue(attrName) && !this.firstMergedCell);
  }

  /**
   * Get the cell note if it has one.
   * 
   * @return {string}
   */
  getTitle() {
    const title = this.getCellAttrValue("cellNote").length > 0 ? `title="${this.getCellAttrValue("cellNote")}"` : '';
    return title;
  }

  /**
   * Get the spreadsheet cell alignment.
   * 
   * @return {string}
   */
  getCellAlignmentStyle() {
    const attrName = "sheetCellType";
    const cellType = this.getCellAttrValue(attrName);
    const cellAlignmentStyle = cellType === 'string' ? "text-align:left" : "text-align:right";
    return cellAlignmentStyle
  }

  /**
   * Check for a number of range cell stylings and build a CSS element for the HTML table style.
   * 
   * @return {string}
   */
  getCellStyle() {
    const backgroundColor = `background-color:${this.getCellAttrValue("backgroundColor")}`;
    const color = `color: ${this.getCellAttrValue("color")}`;
    const fontStyle = `font-style: ${this.getCellAttrValue("fontStyle")}`;
    const fontWeight = `font-weight: ${this.getCellAttrValue("fontWeight")}`;
    const cellAlignmentStyle = this.getCellAlignmentStyle();
    const cellBorder = "border: 1px solid black";
    const cellStyles = [backgroundColor, color, fontStyle, fontWeight, cellAlignmentStyle, cellBorder];
    const cellStyle = 'style="' + cellStyles.join(";") + '"';
    return cellStyle;
  }

  /**
   * Create a formatted HTML table cell.
   * 
   * @todo Examine the logic here and describe it further.
   * 
   * @return {string}
   */
  formatCell() {
    if (this.isMergedToSpan()) {
      return '';
    }
    const rowColSpan = this.isFirstMergedCell() ? this.getRowAndColSpanAttrs() : '';
    const openTag = `<${this.getCellAttrValue("htmlCellType")}`;
    const cellValue = `>${this.getCellAttrValue("value")}`;
    const closeTag = `</${this.getCellAttrValue("htmlCellType")}>`;
    const cellStyle = this.getCellStyle();
    const title = this.getTitle()
    return [openTag, cellStyle, rowColSpan, title, cellValue, closeTag].join(" ");
  }

  /**
   * Create the row and column spans fror HTML cells where the source range cell is merged.
   * 
   * @return {string}
   */
  getRowAndColSpanAttrs() {
    const mergedInfo = this.getCellAttrValue("mergedInfo");
    const { rowSpanCount, colSpanCount } = mergedInfo;
    const rowspan = rowSpanCount > 1 ? `rowspan="${rowSpanCount}"` : '';
    const colspan = colSpanCount > 1 ? `colspan="${colSpanCount}"` : '';
    return rowspan + " " + colspan;
  }

  /**
   * Create a summary object for a formatted cell for use in debugging.
   * 
   * @return {object}
   */
  toString() {
    const me = {
      "tableCellType": this.tableCellType,
      "isFirstMergedCell": this.firstMergedCell,
      "mergedToSpan": this.mergedToSpan,
      "formattedCell": this.formattedCell
    };
    return me;
  }

}

/**
 * Tester function for class CellTransform.
 * 
 * @return {void}
 */
function runCellTransform() {
  // style="background-color: #ff0000;"
  // https://www.lifewire.com/change-table-background-color-3469869
  // https://www.quora.com/How-do-I-add-multiple-CSS-styles-to-a-single-HTML-element
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HTML");
  const rng = sheet.getDataRange();
  const rngCellAttrs = new RangeCellAttributes(rng);
  const cellAttrsObjList = rngCellAttrs.getRngCellAttributes();
  const mergedRngAttrsMap = rngCellAttrs.getMapOfMergedRngAttributes();
  const primaryCellAttrObjID = "address";
  const primaryMergedRnfAttrObjID = "firstCellAddress";
  const attrObjListForHtml = new AttrObjListForHtmlTableMaker(cellAttrsObjList,
    mergedRngAttrsMap,
    primaryCellAttrObjID,
    primaryMergedRnfAttrObjID);

  const mergedInfoAddedToCellAttrsObjList = attrObjListForHtml.addMergedInfoToCellAttrsObjList();
  const lookupKeys = {
    "address": "required",
    "rangeRowIndex": "required",
    "htmlCellType": "required",
    "rangeColIndex": "required",
    "sheetCellType": "required",
    "value": "required",
    "backgroundColor": "required",
    "color": "required",
    "isMerged": "required",
    "cellNote": "required",
    "fontStyle": "required",
    "fontWeight": "required",
    "mergedInfo": "optional"
  };
  const testCell = mergedInfoAddedToCellAttrsObjList[0][3];
  console.log(testCell);
  const cellTrans = new CellTransform(testCell, lookupKeys);
  console.log(cellTrans.toString());
  console.log(cellTrans.formatCell());
}

/**
 * This is the class to be used by clients to generate HTML tables that try, to the greatest extent possible,
 * to replicate the contents and appearance of the Google Sheets range they are taken from.
 * Currently, assumes that the first row of the given range is to be converted to HTML table td elements and the
 * that the remaining rows correspond to tr HTML table rows.
 */
class HtmlGenerator {
  /**
   * Sets up the instance variables for the given Range object.
   * The instance variable contains hard-coded cell attributes that are required.
   * 
   * @param {Range} rng - A Google Sheets Range object containing the cells and formatting to be converted into an HTML table.
   */
  constructor(rng) {
    this.rng = rng;
    this.rngCellAttrs = new RangeCellAttributes(this.rng);
    this.cellAttrsObjList = this.rngCellAttrs.getRngCellAttributes();
    this.mergedRngAttrsMap = this.rngCellAttrs.getMapOfMergedRngAttributes();
    this.lookupKeys = {
      "address": "required",
      "rangeRowIndex": "required",
      "htmlCellType": "required",
      "rangeColIndex": "required",
      "sheetCellType": "required",
      "value": "required",
      "backgroundColor": "required",
      "color": "required",
      "isMerged": "required",
      "cellNote": "required",
      "fontStyle": "required",
      "fontWeight": "required",
      "mergedInfo": "optional"
    };
    this.rowCount = this.cellAttrsObjList.length;
    this.colCount = this.cellAttrsObjList[0].length;
    this.mergedInfoAddedToCellAttrsObjList = this.createCellAttrsObjList();
    this.tableCells = this.getTableCells();
  }

  /**
   * Get the array of arrays where the inner array elements are the cell attributes that need to be translated
   * into HTML table CSS styles and table cell HTML attributes.
   * 
   * @return @return {array} - All the range cells attributes included merged range information.
   */
  createCellAttrsObjList() {
    const primaryCellAttrObjID = "address";
    const primaryMergedRnfAttrObjID = "firstCellAddress";
    const attrObjListForHtml = new AttrObjListForHtmlTableMaker(this.cellAttrsObjList,
      this.mergedRngAttrsMap,
      primaryCellAttrObjID,
      primaryMergedRnfAttrObjID);

    const mergedInfoAddedToCellAttrsObjList = attrObjListForHtml.addMergedInfoToCellAttrsObjList();
    return mergedInfoAddedToCellAttrsObjList;
  }

  /**
   * Get the formatted HTML table cells array of arrays.
   * 
   * @return {array<array>} 
   */
  getTableCells() {
    let tableCells = [];
    for (let i = 0; i < this.rowCount; i++) {
      let row = [];
      for (let j = 0; j < this.colCount; j++) {
        let cell = this.mergedInfoAddedToCellAttrsObjList[i][j];
        let cellTransformed = new CellTransform(cell, this.lookupKeys);
        row.push(cellTransformed.formatCell());
      }
      tableCells.push(row);
    }
    return tableCells;
  }

  /**
   * Create the pre-table part of the HTML page.
   * 
   * @return {string}
   */
  getPageTop() {
    const pageTop = `
<!DOCTYPE html>
<html>
<head>
</head>
<body>
`;
    return pageTop;
  }

  /**
   * Create the post-table part of the HTML page.
   */
  getPageBottom() {
    const pageBottom = `
</body>
</html>  
    `;
    return pageBottom;
  }

  /**
   * Convert the array of arrays of table cells into a HTML table format.
   * 
   * @return {string} - HTML table string
   */
  get TableHtml() {
    const concatedRows = this.TableCells.map((row) => { return "<tr>" + row.join("") + "</tr>"; });
    const tableHtml = '<table style="border-collapse: collapse;">\n' + concatedRows.join("\n") + '\n</table>';
    return tableHtml;
  }

  /**
   * Create a complete HTML webpage containing the generated HTML table.
   * 
   * @return {string}
   */
  get FullPageHtml() {
    let fullPage = this.getPageTop() + this.TableHtml + this.getPageBottom();
    return fullPage;
  }

  /**
   *  Return the array of arrays representing formatted HTML table cells.
   * 
   * @return {array}  
   */
  get TableCells() {
    return this.tableCells;
  }

}

/**
 * Test function for class HtmlGenerator.
 * 
 * @return {void}
 */
function runHtmlGenerator() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HTML");
  const rng = sheet.getDataRange();
  const htmlGen = new HtmlGenerator(rng);
  const tableHtml = htmlGen.FullPageHtml;
  const fileName = "emp_full.html";
  newFile = DriveApp.createFile(fileName, tableHtml);
}
