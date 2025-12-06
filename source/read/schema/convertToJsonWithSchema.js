export default function convertToJsonWithSchema(result, schema) {
  const { rows, errors } = result
  const mappedRows = []

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i]

    if (row === null) {
      continue
    }

    const sheetRowNumber = i + 2

    const mappedRow = {}

    for (const key of Object.keys(schema)) {
      const schemaEntry = schema[key]
      const columnTitle = schemaEntry.column

      const value = (row && row[key] !== undefined) ? row[key] : null

      const cellErrors = errors.filter(error => error.row === sheetRowNumber && error.column === columnTitle)

      let mappedErrors = null
      if (cellErrors.length > 0) {
        mappedErrors = cellErrors.map(error => error.error)
      }

      mappedRow[key] = {
        key: key,
        value: value,
        errors: mappedErrors
      }
    }

    mappedRows.push(mappedRow)
  }

  return {
    rows: mappedRows
  }
}
