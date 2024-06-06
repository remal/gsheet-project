class DefaultFormulas extends AbstractIssueLogic {

    static insertDefaultFormulas(range: Range) {
        const processedRange = this._processRange(range)
        if (processedRange == null) {
            return
        } else {
            range = processedRange
        }

        const sheet = range.getSheet()
        const startRow = range.getRow()
        const endRow = startRow + range.getNumRows() - 1

        const {issues, childIssues} = this._getIssueValues(range)

        const generateAndInsertFormulas = (
            column: number,
            formulaGenerator: (row: number) => string,
        ) => {
            const values = this._getStringValues(range, column)
            const formulas = this._getFormulas(range, column)
            for (let row = startRow; row <= endRow; ++row) {
                const index = row - startRow
                if (!issues[index].length && !childIssues[index].length) {
                    if (formulas[index].length) {
                        sheet.getRange(row, column).setFormula('')
                    }
                    continue
                }

                if (!values[index].length && !formulas[index].length) {
                    const formula = Utils.processFormula(formulaGenerator(row))
                    sheet.getRange(row, column).setFormula(formula)
                }
            }
        }


        const estimateColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName)
        const startColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.startColumnName)
        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName)

        generateAndInsertFormulas(endColumn, row => {
            const startA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, startColumn))
            const estimateA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, estimateColumn))
            return `
                =IF(
                    OR(
                        ISBLANK(${startA1Notation}),
                        ISBLANK(${estimateA1Notation})
                    ),
                    "",
                    WORKDAY(${startA1Notation}, ${estimateA1Notation})
                )
            `
        })
    }

}