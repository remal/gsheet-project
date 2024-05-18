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
                const parentIssueIds = allParentIssueIds[index - 1]
                if (!parentIssueIds?.length) {
                    continue
                }

                const previousParentIssueIds = index >= 2 ? allParentIssueIds[index - 2] : []
                if (Utils.arrayEquals(parentIssueIds, previousParentIssueIds)) {
                    continue
                }

                const issueIndex = 1 + allIssueIds.findIndex(ids =>
                    ids?.some(id => parentIssueIds.includes(id)),
                )
                const newIndex = issueIndex + 1
                if (newIndex === index) {
                    continue
                }

                const newRow = GSheetProjectSettings.firstDataRow + newIndex
                sheet.moveRows(sheet.getRange(index, 1), newRow)
                isChanged = true
                index = Math.min(index, newIndex)
            }

            if (!isChanged) {
                break
            }
        }
    }

}
