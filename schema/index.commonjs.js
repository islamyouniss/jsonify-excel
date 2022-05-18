// This file is deprecated.
// It's the same as `index.cjs`, just retains the old file name.
// Someone might have imported this module as `read-excel-file/schema/index.commonjs`.

exports = module.exports = require('../commonjs/read/schema/convertToJson.js').default
exports['default'] = require('../commonjs/read/schema/convertToJson.js').default