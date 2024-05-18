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
                .map(cols => cols[0].toString())
                .map(text => GSheetProjectSettings.issueIdsExtractor(text))
        }

        // group children:
        grouping: do {
            const allIssueIds = getAllIds(issueIdColumn)
            const allParentIssueIds = getAllIds(parentIssueIdColumn)
            for (let index = allParentIssueIds.length - 1; 0 <= index; --index) {
                const parentIssueIds = allParentIssueIds[index]
                if (!parentIssueIds?.length) {
                    continue
                }

                let previousIndex = null
                for (let prevIndex = index - 1; 0 <= prevIndex; --prevIndex) {
                    const prevParentIssueIds = allParentIssueIds[prevIndex]
                    if (Utils.arrayEquals(parentIssueIds, prevParentIssueIds)) {
                        previousIndex = prevIndex
                    }
                }

                if (previousIndex != null && previousIndex < index - 1) {
                    const newIndex = previousIndex + 1
                    const row = GSheetProjectSettings.firstDataRow + index
                    const newRow = GSheetProjectSettings.firstDataRow + newIndex
                    sheet.moveRows(sheet.getRange(row, 1), newRow)
                    continue grouping;
                }
            }
        } while (false)

        // move children:
        moving: do {
            const allIssueIds = getAllIds(issueIdColumn)
            const allParentIssueIds = getAllIds(parentIssueIdColumn)
            for (let index = 0; index < allParentIssueIds.length; ++index) {
                const currentIndex = index
                const parentIssueIds = allParentIssueIds[currentIndex]
                if (!parentIssueIds?.length) {
                    continue
                }

                let groupSize = 1
                for (; index < allParentIssueIds.length; ++index) {
                    const nextParentIssueIds = allParentIssueIds[index]
                    if (Utils.arrayEquals(parentIssueIds, nextParentIssueIds)) {
                        ++groupSize
                    } else {
                        break
                    }
                }

                const issueIndex = allIssueIds.findIndex(ids =>
                    ids?.some(id => parentIssueIds.includes(id)),
                )
                if (issueIndex < 0 || issueIndex == currentIndex || issueIndex == currentIndex - 1) {
                    continue
                }

                const newIndex = issueIndex + 1
                const row = GSheetProjectSettings.firstDataRow + index
                const newRow = GSheetProjectSettings.firstDataRow + newIndex
                sheet.moveRows(sheet.getRange(row, 1), newRow)
                continue moving;
            }
        } while (false)
    }

}
