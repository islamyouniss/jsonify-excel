// This file is deprecated.
// It's the same as `index.cjs`, just retains the old file name.
// Someone might have imported this module as `read-excel-file/node/index.commonjs`.
// https://gitlab.com/catamphetamine/read-excel-file/-/issues/3#note_896833610
//
// It also fixes the issues when some software doesn't see files with `*.cjs` file extensions
// when used as the `main` property value in `package.json`.

exports = module.exports = require('../commonjs/read/readXlsxFileNode.js').default
exports['default'] = require('../commonjs/read/readXlsxFileNode.js').default
exports.readSheetNames = require('../commonjs/read/readSheetNamesNode.js').default
exports.parseExcelDate = require('../commonjs/read/parseDate.js').default
exports.Integer = require('../commonjs/types/Integer.js').default
exports.Email = require('../commonjs/types/Email.js').default
exports.URL = require('../commonjs/types/URL.js').default