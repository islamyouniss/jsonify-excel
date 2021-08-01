import parseCellValue from './parseCellValue'
import dropEmptyRows from './dropEmptyRows'
import dropEmptyColumns from './dropEmptyColumns'

import {
  getSharedStrings,
  getCellValue,
  getCellInlineStringValue,
  getCells,
  getDimensions,
  getBaseStyles,
  getCellStyles,
  getNumberFormats,
  getWorkbookProperties,
  getRelationships,
  getSheets
} from '../xml/xlsx'

// Maps "A1"-like coordinates to `{ row, column }` numeric coordinates.
const letters = ["", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

// "The minimum viable XLSX reader"
// https://www.brendanlong.com/the-minimum-viable-xlsx-reader.html

/**
 * Reads an (unzipped) XLSX file structure into a 2D array of cells.
 * @param  {object} contents - A list of XML files inside XLSX file (which is a zipped directory).
 * @param  {number?} options.sheet - Workbook sheet id (`1` by default).
 * @param  {string?} options.dateFormat - Date format, e.g. "mm/dd/yyyy". Values having this format template set will be parsed as dates.
 * @param  {object} contents - A list of XML files inside XLSX file (which is a zipped directory).
 * @return {object} An object of shape `{ data, cells, properties }`. `data: string[][]` is an array of rows, each row being an array of cell values. `cells: string[][]` is an array of rows, each row being an array of cells. `properties: object` is the spreadsheet properties (e.g. whether date epoch is 1904 instead of 1900).
 */
export default function readXlsx(contents, xml, options = {}) {
  if (!options.sheet) {
    options = {
      sheet: 1,
      ...options
    }
  }

  // Some Excel editors don't want to use standard naming scheme for sheet files.
  // https://github.com/tidyverse/readxl/issues/104
  const filePaths = parseFilePaths(contents['xl/_rels/workbook.xml.rels'], xml)

  // Default file path for "shared strings": "xl/sharedStrings.xml".
  const values = filePaths.sharedStrings
    ? parseValues(contents[filePaths.sharedStrings], xml)
    : []

  // Default file path for "styles": "xl/styles.xml".
  const styles = filePaths.styles
    ? parseStyles(contents[filePaths.styles], xml)
    : {}

  const properties = parseProperties(contents['xl/workbook.xml'], xml)

  // A feature for getting the list of sheets in an Excel file.
  // https://github.com/catamphetamine/read-excel-file/issues/14
  if (options.getSheets) {
    return properties.sheets.map(({ name }) => ({
      name
    }))
  }

  // Find the sheet by name, or take the first one.
  let sheetRelationId
  if (typeof options.sheet === 'number') {
    const _sheet = properties.sheets[options.sheet - 1]
    sheetRelationId = _sheet && _sheet.relationId
  } else {
    for (const sheet of properties.sheets) {
      if (sheet.name === options.sheet) {
        sheetRelationId = sheet.relationId
        break
      }
    }
  }

  // If the sheet wasn't found then throw an error.
  // Example: "xl/worksheets/sheet1.xml".
  if (!sheetRelationId || !filePaths.sheets[sheetRelationId]) {
    throw createSheetNotFoundError(options.sheet, properties.sheets)
  }

  // Parse sheet data.
  const sheet = parseSheet(
    contents[filePaths.sheets[sheetRelationId]],
    xml,
    values,
    styles,
    properties,
    options
  )

  // If the sheet is empty.
  if (sheet.cells.length === 0) {
    if (options.properties) {
      return {
        data: [],
        properties
      }
    }
    return []
  }

  const [ leftTop, rightBottom ] = sheet.dimensions

  const colsCount = (rightBottom.column - leftTop.column) + 1
  const rowsCount = (rightBottom.row - leftTop.row) + 1

  // `sheet.cells` seem to not necessarily be sorted by row and column.
  let data = new Array(rowsCount)
  let i = 0
  while (i < rowsCount) {
    data[i] = new Array(colsCount)
    let j = 0
    while (j < colsCount) {
      data[i][j] = null
      j++
    }
    i++
  }

  for (const cell of sheet.cells) {
    const row = cell.row - leftTop.row
    const column = cell.column - leftTop.column
    data[row][column] = cell.value
  }

  // Fill in the row map.
  const { rowMap } = options
  if (rowMap) {
    let i = 0
    while (i < data.length) {
      rowMap[i] = i
      i++
    }
  }

  data = dropEmptyRows(
    dropEmptyColumns(data, { onlyTrimAtTheEnd: true }),
    { onlyTrimAtTheEnd: true, rowMap }
  )

  if (options.transformData) {
    data = options.transformData(data)
    // data = options.transformData(data, {
    //   dropEmptyRowsAndColumns(data) {
    //     return dropEmptyRows(dropEmptyColumns(data), { rowMap })
    //   }
    // })
  }

  if (options.properties) {
    return {
      data,
      properties
    }
  }

  return data
}

function calculateDimensions (cells) {
  const comparator = (a, b) => a - b
  const allRows = cells.map(cell => cell.row).sort(comparator)
  const allCols = cells.map(cell => cell.column).sort(comparator)
  const minRow = allRows[0]
  const maxRow = allRows[allRows.length - 1]
  const minCol = allCols[0]
  const maxCol = allCols[allCols.length - 1]

  return [
    { row: minRow, column: minCol },
    { row: maxRow, column: maxCol }
  ]
}

function columnLetterToNumber(col) {
  // `for ... of ...` would require Babel polyfill for iterating a string.
  let n = 0
  let i = 0
  while (i < col.length) {
    n *= 26
    n += letters.indexOf(col[i])
    i++
  }
  return n
}

function parseCellCoordinates(coords) {
  // Examples: "AA2091", "R988", "B1"
  coords = coords.split(/(\d+)/)
  return [
    // Row.
    parseInt(coords[1]),
    // Column.
    columnLetterToNumber(coords[0].trim())
  ]
}

// Example of a `<c/>`ell element:
//
// <c>
//    <f>string</f> — formula.
//    <v>string</v> — formula pre-computed value.
//    <is>
//       <t>string</t> — an `inlineStr` string (rather than a "common string" from a dictionary).
//       <r>
//          <rPr>
//            ...
//          </rPr>
//          <t>string</t>
//       </r>
//       <rPh sb="1" eb="1">
//          <t>string</t>
//       </rPh>
//       <phoneticPr fontId="1"/>
//    </is>
//    <extLst>
//       <ext>
//          <!--any element-->
//       </ext>
//    </extLst>
// </c>
//
function parseCell(cellNode, sheet, xml, values, styles, properties, options) {
  const coords = parseCellCoordinates(cellNode.getAttribute('r'))

  const valueElement = getCellValue(sheet, cellNode)

  // For `xpath`, `value` can be `undefined` while for native `DOMParser` it's `null`.
  // So using `value && ...` instead of `if (value !== undefined) { ... }` here
  // for uniform compatibility with both `xpath` and native `DOMParser`.
  let value = valueElement && valueElement.textContent

  let type
  if (cellNode.hasAttribute('t')) {
    type = cellNode.getAttribute('t')
  }

  return {
    row: coords[0],
    column: coords[1],
    value: parseCellValue(value, type, {
      getInlineStringValue: () => getCellInlineStringValue(cellNode),
      getStyleId: () => cellNode.getAttribute('s'),
      styles,
      values,
      properties,
      options
    })
  }
}

function parseSheet(content, xml, values, styles, properties, options) {
  const sheet = xml.createDocument(content)

  let cells = getCells(sheet)

  if (cells.length === 0) {
    return { cells: [] }
  }

  cells = cells.map((node) => {
    return parseCell(node, sheet, xml, values, styles, properties, options)
  })

  let dimensions = getDimensions(sheet)
  if (dimensions) {
    dimensions = dimensions.split(':').map(parseCellCoordinates).map(([row, column]) => ({
      row,
      column
    }))
    // When there's only a single cell on a sheet
    // there can sometimes be just "A1" for the dimensions string.
    if (dimensions.length === 1) {
      dimensions = [dimensions[0], dimensions[0]]
    }
  } else {
    dimensions = calculateDimensions(cells)
  }

  return { cells, dimensions }
}

function parseValues(content, xml) {
  if (!content) {
    return []
  }
  return getSharedStrings(xml.createDocument(content))
}

// http://officeopenxml.com/SSstyles.php
// Returns an array of cell styles.
// A cell style index is its ID.
function parseStyles(content, xml) {
  if (!content) {
    return {}
  }

  // https://social.msdn.microsoft.com/Forums/sqlserver/en-US/708978af-b598-45c4-a598-d3518a5a09f0/howwhen-is-cellstylexfs-vs-cellxfs-applied-to-a-cell?forum=os_binaryfile
  // https://www.office-forums.com/threads/cellxfs-cellstylexfs.2163519/
  const doc = xml.createDocument(content)

  const baseStyles = getBaseStyles(doc)
    .map(parseCellStyle)

  const numberFormats = getNumberFormats(doc)
    .map(parseNumberFormatStyle)
    .reduce((formats, format) => {
      // Format ID is a numeric index.
      // There're some standard "built-in" formats (in Excel) up to about `100`.
      formats[format.id] = format
      return formats
    }, [])

  const getCellStyle = (xf) => {
    if (xf.hasAttribute('xfId')) {
      return {
        ...baseStyles[xf.xfId],
        ...parseCellStyle(xf, numberFormats)
      }
    }
    return parseCellStyle(xf, numberFormats)
  }

  return getCellStyles(doc).map(getCellStyle)
}

function parseNumberFormatStyle(numFmt) {
  return {
    id: numFmt.getAttribute('numFmtId'),
    template: numFmt.getAttribute('formatCode')
  }
}

// http://www.datypic.com/sc/ooxml/e-ssml_xf-2.html
function parseCellStyle(xf, numFmts) {
  const style = {}
  if (xf.hasAttribute('numFmtId')) {
    const numberFormatId = xf.getAttribute('numFmtId')
    // Built-in number formats don't have a `<numFmt/>` element in `styles.xml`.
    // https://hexdocs.pm/xlsxir/number_styles.html
    if (numFmts[numberFormatId]) {
      style.numberFormat = numFmts[numberFormatId]
    } else {
      style.numberFormat = { id: numberFormatId }
    }
  }
  return style
}

// I guess `xl/workbook.xml` file should always be present inside the *.xlsx archive.
function parseProperties(content, xml) {
  const book = xml.createDocument(content)

  const properties = {};

  // Read `<workbookPr/>` element to detect whether dates are 1900-based or 1904-based.
  // https://support.microsoft.com/en-gb/help/214330/differences-between-the-1900-and-the-1904-date-system-in-excel
  // http://webapp.docx4java.org/OnlineDemo/ecma376/SpreadsheetML/workbookPr.html

  const workbookProperties = getWorkbookProperties(book)

  if (workbookProperties && workbookProperties.getAttribute('date1904') === '1') {
    properties.epoch1904 = true
  }

  // Get sheets info (indexes, names, if they're available).
  // Example:
  // <sheets>
  //   <sheet
  //     xmlns:ns="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  //     name="Sheet1"
  //     sheetId="1"
  //     ns:id="rId3"/>
  // </sheets>
  // http://www.datypic.com/sc/ooxml/e-ssml_sheet-1.html

  properties.sheets = []

  const addSheetInfo = (sheet) => {
    if (sheet.getAttribute('name')) {
      properties.sheets.push({
        id: sheet.getAttribute('sheetId'),
        name: sheet.getAttribute('name'),
        relationId: sheet.getAttribute('r:id')
      })
    }
  }

  getSheets(book).forEach(addSheetInfo)

  return properties;
}

/**
 * Returns sheet file paths.
 * Seems that the correct place to look for the `sheetId` -> `filename` mapping
 * is `xl/_rels/workbook.xml.rels` file.
 * https://github.com/tidyverse/readxl/issues/104
 * @param  {string} content — `xl/_rels/workbook.xml.rels` file contents.
 * @param  {object} xml
 * @return {object}
 */
function parseFilePaths(content, xml) {
  // Example:
  // <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  //   ...
  //   <Relationship
  //     Id="rId3"
  //     Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  //     Target="worksheets/sheet1.xml"/>
  // </Relationships>
  const document = xml.createDocument(content)

  const filePaths = {
    sheets: {},
    sharedStrings: undefined,
    styles: undefined
  }

  const addFilePathInfo = (relationship) => {
    const filePath = relationship.getAttribute('Target')
    const fileType = relationship.getAttribute('Type')
    switch (fileType) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles':
        filePaths.styles = getFilePath(filePath)
        break
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings':
        filePaths.sharedStrings = getFilePath(filePath)
        break
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet':
        filePaths.sheets[relationship.getAttribute('Id')] = getFilePath(filePath)
        break
    }
  }

  getRelationships(document).forEach(addFilePathInfo)

  // Seems like "sharedStrings.xml" is not required to exist.
  // For example, when the spreadsheet doesn't contain any strings.
  // https://github.com/catamphetamine/read-excel-file/issues/85
  // if (!filePaths.sharedStrings) {
  //   throw new Error('"sharedStrings.xml" file not found in the *.xlsx file')
  // }

  return filePaths
}

function getFilePath(path) {
  // Normally, `path` is a relative path inside the ZIP archive,
  // like "worksheets/sheet1.xml", or "sharedStrings.xml", or "styles.xml".
  // There has been one weird case when file path was an absolute path,
  // like "/xl/worksheets/sheet1.xml" (specifically for sheets):
  // https://github.com/catamphetamine/read-excel-file/pull/95
  // Other libraries (like `xlsx`) and software (like Google Docs)
  // seem to support such absolute file paths, so this library does too.
  if (path[0] === '/') {
    return path.slice('/'.length)
  }
  // // Seems like a path could also be a URL.
  // // http://officeopenxml.com/anatomyofOOXML-xlsx.php
  // if (/^[a-z]+\:\/\//.test(path)) {
  //   return path
  // }
  return 'xl/' + path
}

function createSheetNotFoundError(sheet, sheets) {
  const sheetsList = sheets && sheets.map((sheet, i) => `"${sheet.name}" (#${i + 1})`).join(', ')
  return new Error(`Sheet ${typeof sheet === 'number' ? '#' + sheet : '"' + sheet + '"'} not found in the *.xlsx file.${sheets ? ' Available sheets: ' + sheetsList + '.' : ''}`)
}