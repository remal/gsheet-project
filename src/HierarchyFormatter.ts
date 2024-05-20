class HierarchyFormatter {

    static formatHierarchy(range: Range) {
        if (!RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.issueIdColumnName)
            && !RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.parentIssueIdColumnName)
        ) {
            return
        }

        this._formatSheetHierarchy(range.getSheet())
    }

    static formatAllHierarchy() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            this._formatSheetHierarchy(sheet)
        }
    }

    private static _formatSheetHierarchy(sheet: Sheet) {
        if (State.isStructureChanged()) return

        const issueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.issueIdColumnName)
        const parentIssueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.parentIssueIdColumnName)
        if (issueIdColumn == null || parentIssueIdColumn == null) {
            return
        }

        this._groupChildren(sheet)
        this._moveChildren(sheet)
        this._updateTimelineTitleFormula(sheet)
        this._updateDeadlineFormula(sheet)
    }

    private static _groupChildren(sheet: Sheet) {
        if (State.isStructureChanged()) return

        const parentIssueIdColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.parentIssueIdColumnName)

        while (true) {
            const allParentIssueIds = this._getAllIds(sheet, parentIssueIdColumn)
            if (allParentIssueIds.every(ids => !ids?.length)) {
                return
            }

            let isChanged = false
            for (let index = allParentIssueIds.length - 1; 0 <= index; --index) {
                const parentIssueIds = allParentIssueIds[index]
                if (!parentIssueIds?.length) {
                    continue
                }

                let previousIndex: number | null = null
                for (let prevIndex = index - 1; 0 <= prevIndex; --prevIndex) {
                    const prevParentIssueIds = allParentIssueIds[prevIndex]
                    if (Utils.arrayEquals(parentIssueIds, prevParentIssueIds)) {
                        previousIndex = prevIndex
                        break
                    }
                }

                if (previousIndex != null && previousIndex < index - 1) {
                    if (State.isStructureChanged()) return
                    const newIndex = previousIndex + 1
                    const row = GSheetProjectSettings.firstDataRow + index
                    const newRow = GSheetProjectSettings.firstDataRow + newIndex
                    sheet.moveRows(sheet.getRange(row, 1), newRow)
                    isChanged = true
                }
            }

            if (!isChanged) {
                break
            }
        }
    }

    private static _moveChildren(sheet: Sheet) {
        if (State.isStructureChanged()) return

        const issueIdColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueIdColumnName)
        const parentIssueIdColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.parentIssueIdColumnName)

        while (true) {
            const allIssueIds = this._getAllIds(sheet, issueIdColumn)
            const allParentIssueIds = this._getAllIds(sheet, parentIssueIdColumn)
            let isChanged = false
            for (let index = 0; index < allParentIssueIds.length; ++index) {
                const currentIndex = index
                const parentIssueIds = allParentIssueIds[currentIndex]
                if (!parentIssueIds?.length) {
                    continue
                }

                let groupSize = 1
                for (index; index < allParentIssueIds.length - 1; ++index) {
                    const nextParentIssueIds = allParentIssueIds[index + 1]
                    if (Utils.arrayEquals(parentIssueIds, nextParentIssueIds)) {
                        ++groupSize
                    } else {
                        break
                    }
                }

                const issueIndex = allIssueIds.findIndex((ids, issueIndex) =>
                    ids?.some(id => parentIssueIds.includes(id))
                    && issueIndex !== currentIndex,
                )
                if (issueIndex < 0 || issueIndex == currentIndex - 1) {
                    continue
                }

                if (State.isStructureChanged()) return
                const newIndex = issueIndex + 1
                const row = GSheetProjectSettings.firstDataRow + currentIndex
                const newRow = GSheetProjectSettings.firstDataRow + newIndex
                sheet.moveRows(sheet.getRange(row, 1, groupSize, 1), newRow)
                break
            }

            if (!isChanged) {
                break
            }
        }
    }

    private static _updateTimelineTitleFormula(sheet: Sheet) {
        if (State.isStructureChanged()) return

        const timelineTitleColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.timelineTitleColumnName)
        const titleColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.titleColumnName)
        if (timelineTitleColumn == null || titleColumn == null) {
            return
        }

        const issueIdColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueIdColumnName)
        const parentIssueIdColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.parentIssueIdColumnName)

        const allIssueIds = this._getAllIds(sheet, issueIdColumn)
        const allParentIssueIds = this._getAllIds(sheet, parentIssueIdColumn)

        const timelineTitleRange = SheetUtils.getColumnRange(
            sheet,
            GSheetProjectSettings.timelineTitleColumnName!,
            GSheetProjectSettings.firstDataRow,
        )
        const timelineTitleFormulas = timelineTitleRange.getFormulas()

        let isChanged = false
        for (let index = 0; index < allParentIssueIds.length; ++index) {
            const row = GSheetProjectSettings.firstDataRow + index
            let formula = `=${sheet.getRange(row, titleColumn).getA1Notation()}`

            const parentIssueIds = allParentIssueIds[index]
            if (parentIssueIds?.length) {
                const issueIndex = allIssueIds.findIndex((ids, issueIndex) =>
                    ids?.some(id => parentIssueIds.includes(id))
                    && issueIndex !== index,
                )
                if (issueIndex >= 0) {
                    const issueRow = GSheetProjectSettings.firstDataRow + issueIndex
                    const formulaCondition = `ISBLANK(${sheet.getRange(row, titleColumn).getA1Notation()})`
                    const formulaTrue = `${sheet.getRange(issueRow, titleColumn).getA1Notation()}`
                    const formulaFalse = `${sheet.getRange(row, titleColumn).getA1Notation()}`
                    formula = `=IF(${formulaCondition}, ${formulaTrue}, ${formulaFalse})`
                }
            }

            if (!Utils.arrayEquals(timelineTitleFormulas[index], [formula])) {
                timelineTitleFormulas[index] = [formula]
                isChanged = true
            }
        }

        if (isChanged) {
            if (State.isStructureChanged()) return
            timelineTitleRange.setFormulas(timelineTitleFormulas)
        }
    }

    private static _updateDeadlineFormula(sheet: Sheet) {
        if (State.isStructureChanged()) return

        const deadlineColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.deadlineColumnName)
        const titleColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.titleColumnName)
        if (timelineTitleColumn == null || titleColumn == null) {
            return
        }

        const issueIdColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueIdColumnName)
        const parentIssueIdColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.parentIssueIdColumnName)

        const allIssueIds = this._getAllIds(sheet, issueIdColumn)
        const allParentIssueIds = this._getAllIds(sheet, parentIssueIdColumn)

        const timelineTitleRange = SheetUtils.getColumnRange(
            sheet,
            GSheetProjectSettings.timelineTitleColumnName!,
            GSheetProjectSettings.firstDataRow,
        )
        const timelineTitleFormulas = timelineTitleRange.getFormulas()

        let isChanged = false
        for (let index = 0; index < allParentIssueIds.length; ++index) {
            const row = GSheetProjectSettings.firstDataRow + index
            let formula = `=${sheet.getRange(row, titleColumn).getA1Notation()}`

            const parentIssueIds = allParentIssueIds[index]
            if (parentIssueIds?.length) {
                const issueIndex = allIssueIds.findIndex((ids, issueIndex) =>
                    ids?.some(id => parentIssueIds.includes(id))
                    && issueIndex !== index,
                )
                if (issueIndex >= 0) {
                    const issueRow = GSheetProjectSettings.firstDataRow + issueIndex
                    const formulaCondition = `ISBLANK(${sheet.getRange(row, titleColumn).getA1Notation()})`
                    const formulaTrue = `${sheet.getRange(issueRow, titleColumn).getA1Notation()}`
                    const formulaFalse = `${sheet.getRange(row, titleColumn).getA1Notation()}`
                    formula = `=IF(${formulaCondition}, ${formulaTrue}, ${formulaFalse})`
                }
            }

            if (!Utils.arrayEquals(timelineTitleFormulas[index], [formula])) {
                timelineTitleFormulas[index] = [formula]
                isChanged = true
            }
        }

        if (isChanged) {
            if (State.isStructureChanged()) return
            timelineTitleRange.setFormulas(timelineTitleFormulas)
        }
    }

    private static _getAllIds(sheet: Sheet, column: number | string): (string[] | null)[] {
        return SheetUtils.getColumnRange(sheet, column, GSheetProjectSettings.firstDataRow)
            .getValues()
            .map(cols => cols[0].toString())
            .map(text => GSheetProjectSettings.issueIdsExtractor(text))
    }

}
