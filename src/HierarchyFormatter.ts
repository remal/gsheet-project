class HierarchyFormatter {

    static formatHierarchy(range: Range) {
        if (!RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.issueIdColumnName)
            && !RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.parentIssueIdColumnName)
        ) {
            return
        }

        this.formatSheetHierarchy(range.getSheet())
    }

    static formatAllHierarchy() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            this.formatSheetHierarchy(sheet)
        }
    }

    private static formatSheetHierarchy(sheet: Sheet) {
        const issueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.issueIdColumnName)
        const parentIssueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.parentIssueIdColumnName)
        if (issueIdColumn == null || parentIssueIdColumn == null) {
            return
        }

        const lastRow = sheet.getLastRow()
        const getAllIds = (column: number): (string[] | null)[] => {
            return sheet.getRange(GSheetProjectSettings.firstDataRow, column, lastRow, 1)
                .getValues()
                .map(it => it[0].toString())
                .map(GSheetProjectSettings.issueIdsExtractor)
        }

        while (true) {
            const allIssueIds = getAllIds(issueIdColumn)
            const allParentIssueIds = getAllIds(parentIssueIdColumn)
            let isChanged = false
            for (let index = allParentIssueIds.length - 1; 0 <= index; --index) {
                const parentIssueIds = allParentIssueIds[index]
                if (!parentIssueIds?.length) {
                    continue
                }

                const previousParentIssueIds = index >= 1 ? allParentIssueIds[index - 1] : []
                if (Utils.arrayEquals(parentIssueIds, previousParentIssueIds)) {
                    continue
                }

                const issueIndex = allIssueIds.findIndex(ids =>
                    ids?.some(id => parentIssueIds.includes(id)),
                )
                if (issueIndex == null || issueIndex === index) {
                    continue
                }

                const newIndex = issueIndex + 1
                if (newIndex === index) {
                    continue
                }

                const row = GSheetProjectSettings.firstDataRow + index
                const newRow = GSheetProjectSettings.firstDataRow + newIndex
                sheet.moveRows(sheet.getRange(row, 1), newRow)
                isChanged = true
                if (newIndex > index) {
                    break
                } else {
                    index = newIndex
                }
            }

            if (!isChanged) {
                break
            }
        }
    }

}
