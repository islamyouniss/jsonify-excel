import convertToJsonWithSchema from './convertToJsonWithSchema.js'

describe('convertToJsonWithSchema', () => {
    it('should convert output to the desired format', () => {
        const input = {
            rows: [
                {
                    trainingType: "program",
                    name: "pro 1",
                    category: "leaders_learning",
                    targetAudience: "pro 1 TA"
                },
                null,
                null
            ],
            errors: [
                {
                    error: "required",
                    row: 2,
                    column: "Start Date"
                },
                {
                    error: "required",
                    row: 2,
                    column: "End Date"
                }
            ]
        }

        const schema = {
            trainingType: {
                column: "Training Type",
                type: String
            },
            name: {
                column: "Name",
                type: String
            },
            category: {
                column: "Category",
                type: String
            },
            targetAudience: {
                column: "Target Audience",
                type: String
            },
            startDate: {
                column: "Start Date",
                type: Date,
                required: true
            },
            endDate: {
                column: "End Date",
                type: Date,
                required: true
            }
        }

        const expected = {
            rows: [
                {
                    trainingType: {
                        value: "program",
                        errors: null
                    },
                    name: {
                        value: "pro 1",
                        errors: null
                    },
                    category: {
                        value: "leaders_learning",
                        errors: null
                    },
                    targetAudience: {
                        value: "pro 1 TA",
                        errors: null
                    },
                    startDate: {
                        value: null,
                        errors: "required"
                    },
                    endDate: {
                        value: null,
                        errors: "required"
                    },
                }
            ]
        }

        const result = convertToJsonWithSchema(input, schema)
        result.should.deep.equal(expected)
    })
})
