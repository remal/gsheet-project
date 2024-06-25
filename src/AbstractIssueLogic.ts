abstract class AbstractIssueLogic {

    protected static _processRange(range: Range): Range | null {
        if (![
            GSheetProjectSettings.issueKeyColumnName,
            GSheetProjectSettings.childIssueKeyColumnName,
            GSheetProjectSettings.teamColumnName,
        ].some(columnName =>
            RangeUtils.doesRangeHaveSheetColumn(range, GSheetProjectSettings.sheetName, columnName),
        )) {
            return null
        }

        const sheet = range.getSheet()
        ProtectionLocks.lockAllColumns(sheet)

        range = RangeUtils.withMinMaxRows(range, GSheetProjectSettings.firstDataRow, SheetUtils.getLastRow(sheet))
        const startRow = range.getRow()
        const rows = range.getNumRows()
        const endRow = startRow + rows - 1
        ProtectionLocks.lockRows(sheet, endRow)
        return range
    }

    protected static _getIssueValues(range: Range): IssueColumnValues {
        const sheet = range.getSheet()
        const startRow = range.getRow()
        const endRow = startRow + range.getNumRows() - 1
        const result = SheetUtils.getColumnsStringValues(sheet, {
            issues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName),
            childIssues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueKeyColumnName),
        }, startRow, endRow)

        Utils.trimArrayEndBy(result.issues, it => !it?.length)
        result.childIssues.length = result.issues.length

        return result
    }

    protected static _getIssueValuesWithLastReloadDate(range: Range): IssueColumnValuesWithLastDataReload {
        const sheet = range.getSheet()
        const startRow = range.getRow()
        const endRow = startRow + range.getNumRows() - 1
        const result = SheetUtils.getColumnsValues(sheet, {
            issues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName),
            childIssues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueKeyColumnName),
            lastDataReload: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.lastDataReloadColumnName),
        }, startRow, endRow)

        Utils.trimArrayEndBy(result.issues, it => !it?.toString()?.length)
        result.childIssues.length = result.issues.length
        result.lastDataReload.length = result.issues.length

        return {
            issues: result.issues.map(it => it?.toString()),
            childIssues: result.childIssues.map(it => it?.toString()),
            lastDataReload: result.lastDataReload.map(it => Utils.parseDate(it)),
        }
    }

    protected static _getValues(range: Range, column: Column): any[] {
        return RangeUtils.toColumnRange(range, column)!.getValues()
            .map(it => it[0])
    }

    protected static _getStringValues(range: Range, column: Column): string[] {
        return this._getValues(range, column).map(it => it.toString())
    }

    protected static _getFormulas(range: Range, column: Column): Formula[] {
        return RangeUtils.toColumnRange(range, column)!.getFormulas()
            .map(it => it[0])
    }


}

interface IssueColumnValues {
    issues: string[]
    childIssues: string[]
}

interface IssueColumnValuesWithLastDataReload extends IssueColumnValues {
    lastDataReload: (Date | null)[]
}
