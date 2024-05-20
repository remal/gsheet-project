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

        const getAllIds = (column: number): (string[] | null)[] => {
            return SheetUtils.getColumnRange(sheet, column, GSheetProjectSettings.firstDataRow)
                .getValues()
                .map(cols => cols[0].toString())
                .map(text => GSheetProjectSettings.issueIdsExtractor(text))
        }

        // group children:
        while (true) {
            const allParentIssueIds = getAllIds(parentIssueIdColumn)
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

        // move children:
        while (true) {
            const allIssueIds = getAllIds(issueIdColumn)
            const allParentIssueIds = getAllIds(parentIssueIdColumn)
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

        // timeline title:
        const timelineTitleColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.timelineTitleColumnName)
        const titleColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.titleColumnName)
        if (timelineTitleColumn != null && titleColumn != null) {
            const allIssueIds = getAllIds(issueIdColumn)
            const allParentIssueIds = getAllIds(parentIssueIdColumn)
            const timelineTitleRange = SheetUtils.getColumnRange(
                sheet,
                GSheetProjectSettings.timelineTitleColumnName!,
                GSheetProjectSettings.firstDataRow,
            )
            const timelineTitleFormulas = timelineTitleRange.getFormulas()

            let isChanged = false
            for (let index = 0; index < allParentIssueIds.length; ++index) {
                const row = GSheetProjectSettings.firstDataRow + index
                const parentIssueIds = allParentIssueIds[index]
                if (!parentIssueIds?.length) {
                    continue
                }

                const issueIndex = allIssueIds.findIndex((ids, issueIndex) =>
                    ids?.some(id => parentIssueIds.includes(id))
                    && issueIndex !== index,
                )
                let formula = `=${sheet.getRange(row, titleColumn).getA1Notation()}`
                if (issueIndex >= 0) {
                    const issueRow = GSheetProjectSettings.firstDataRow + index
                    const formulaCondition = `ISBLANK(${sheet.getRange(row, titleColumn).getA1Notation()})`
                    const formulaTrue = `${sheet.getRange(issueRow, titleColumn).getA1Notation()}`
                    const formulaFalse = `${sheet.getRange(row, titleColumn).getA1Notation()}`
                    formula = `=IF(ISBLANK(${formulaCondition}), ${formulaTrue}, ${formulaFalse})`
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
    }

}

type TimelineTitleFormulaSetter = (childIndex: number, childRow: number, parentRow: number | null) => unknown
