class IssueHierarchyFormatter extends AbstractIssueLogic {

    static formatHierarchy(range: Range) {
        const processedRange = this._processRange(range)
        if (processedRange == null) {
            return
        } else {
            range = processedRange
        }

        const sheet = range.getSheet()
        const startRow = range.getRow()
        const endRow = startRow + range.getNumRows() - 1

        const {issues, childIssues} = this._getIssueValues(sheet.getRange(
            GSheetProjectSettings.firstDataRow,
            range.getColumn(),
            endRow - GSheetProjectSettings.firstDataRow + 1,
            range.getNumColumns(),
        ))

        const issueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName)

        for (let row = startRow; row <= endRow; ++row) {
            const index = row - GSheetProjectSettings.firstDataRow
            const issue = issues[index]
            const childIssue = childIssues[index]
            if (!issue?.length) {
                continue
            }

            const issueRange = sheet.getRange(row, issueColumn)
            if (!childIssue?.length) {
                issueRange.setFontSize(GSheetProjectSettings.fontSize)
                continue
            }

            const parentIssueIndex = issues.indexOf(issue)
            if (parentIssueIndex < 0) {
                continue
            }
            if (childIssues[parentIssueIndex]?.length) {
                continue
            }

            const parentIssueRow = GSheetProjectSettings.firstDataRow + parentIssueIndex
            const parentIssueRange = sheet.getRange(parentIssueRow, issueColumn)
            issueRange
                .setFormula(Formulas.processFormula(`=
                    ${RangeUtils.getAbsoluteA1Notation(parentIssueRange)}
                `))
                .setFontSize(GSheetProjectSettings.fontSize - 2)
        }
    }

}
