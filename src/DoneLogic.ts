/*
class DoneLogic extends AbstractIssueLogic {

    static executeDoneLogic(range: Range) {
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

        const hasIssue = (row: number): boolean => {
            const index = row - startRow
            return !!issues[index]?.length || !!childIssues[index]?.length
        }


        const doneColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.doneColumnName)
        let doneValues = this._getStringValues(range, doneColumn)

        Utils.timed(`checkboxes`, () => {
            const checkboxesA1Notations = Array.from(Utils.range(startRow, endRow))
                .filter(row => hasIssue(row))
                .map(row => sheet.getRange(row, doneColumn).getA1Notation())
            if (checkboxesA1Notations.length) {
                sheet.getRangeList(checkboxesA1Notations).insertCheckboxes()
            }
            const notCheckboxesA1Notations = Array.from(Utils.range(startRow, endRow))
                .filter(row => !hasIssue(row))
                .filter(row => doneValues[row - startRow]?.length)
                .map(row => sheet.getRange(row, doneColumn).getA1Notation())
            if (notCheckboxesA1Notations.length) {
                sheet.getRangeList(notCheckboxesA1Notations).removeCheckboxes().setValue('')
            }

            doneValues = this._getStringValues(range, doneColumn)
        })


        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName)
        const endValues = this._getValues(range, endColumn)

        for (let row = startRow; row <= endRow; ++row) {
            if (!hasIssue(row)) {
                continue
            }

            const index = row - startRow
            let doneValue = doneValues[index].toLowerCase()

            const endRange = sheet.getRange(row, endColumn)
            const rowRange = sheet.getRange(`${row}:${row}`)

            if (doneValue === 'true') {
                const endValue = endValues[index]
                let endDate: Date
                if (Utils.isString(endValue)) {
                    endDate = new Date(Number.isNaN(endValue) ? endValue : parseFloat(endValue))
                } else if (Utils.isNumber(endValue)) {
                    endDate = new Date(endValue)
                } else {
                    try {
                        endDate = new Date(endValue.toString())
                    } catch (e) {
                        console.warn(`Can't get date from ${endRange.getA1Notation()}`)
                        continue
                    }
                }

                if (GSheetProjectSettings.restoreUndoneEnd) {
                    const developerMetadata = rowRange.getDeveloperMetadata()

                    const previousFormulaMetadata = developerMetadata.find(it =>
                        it.getKey() === `${DoneLogic.name}|before-done-end-formula`,
                    )
                    if (previousFormulaMetadata != null) {
                        rowRange.addDeveloperMetadata(
                            `${DoneLogic.name}|before-done-end-formula`,
                            endRange.getFormula(),
                        )
                    }

                    const previousValueMetadata = developerMetadata.find(it =>
                        it.getKey() === `${DoneLogic.name}|before-done-end-value`,
                    )
                    if (previousValueMetadata != null) {
                        rowRange.addDeveloperMetadata(
                            `${DoneLogic.name}|before-done-end-value`,
                            endDate.toString(),
                        )
                    }
                }

                const now = new Date()
                if (now.getTime() < endDate.getTime()) {
                    endRange.setValue(now)
                } else {
                    endRange.setValue(endDate)
                }

            } else if (GSheetProjectSettings.restoreUndoneEnd) {
                const developerMetadata = rowRange.getDeveloperMetadata()
                const previousFormulaMetadata = developerMetadata.find(it =>
                    it.getKey() === `${DoneLogic.name}|before-done-end-formula`,
                )
                const previousValueMetadata = developerMetadata.find(it =>
                    it.getKey() === `${DoneLogic.name}|before-done-end-value`,
                )
                try {
                    const previousFormula = previousFormulaMetadata?.getValue()
                    const previousValue = previousValueMetadata?.getValue()
                    if (previousFormulaMetadata != null && previousFormula?.length) {
                        endRange.setFormula(previousFormula)
                    } else if (previousValueMetadata != null) {
                        if (previousValue?.length) {
                            endRange.setValue(new Date(previousValue))
                        } else {
                            endRange.setValue('')
                        }
                    }
                } finally {
                    previousFormulaMetadata?.remove()
                    previousValueMetadata?.remove()
                }
            }
        }
    }

}
*/
