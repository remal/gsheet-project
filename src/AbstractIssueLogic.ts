abstract class AbstractIssueLogic {

    protected static _processRange(range: Range): Range | null {
        if (![GSheetProjectSettings.issueColumnName, GSheetProjectSettings.titleColumnName].some(columnName =>
            RangeUtils.doesRangeHaveSheetColumn(range, GSheetProjectSettings.sheetName, columnName),
        )) {
            return null
        }

        const sheet = range.getSheet()
        ProtectionLocks.lockAllColumns(sheet)

        range = RangeUtils.withMinMaxRows(range, GSheetProjectSettings.firstDataRow, sheet.getLastRow())
        const startRow = range.getRow()
        const rows = range.getNumRows()
        const endRow = startRow + rows - 1
        ProtectionLocks.lockRows(sheet, endRow)
        return range
    }

    protected static _getIssueValues(range: Range): { issues: string[], childIssues: string[] } {
        const sheet = range.getSheet()
        const startRow = range.getRow()
        const endRow = startRow + range.getNumRows() - 1
        return SheetUtils.getColumnsStringValues(sheet, {
            issues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName),
            childIssues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName),
        }, startRow, endRow)
    }

    protected static _getStringValues(range: Range, column: number): string[] {
        return RangeUtils.toColumnRange(range, column)!.getValues()
            .map(it => it[0].toString())
    }

    protected static _getFormulas(range: Range, column: number): string[] {
        return RangeUtils.toColumnRange(range, column)!.getFormulas()
            .map(it => it[0])
    }


}
