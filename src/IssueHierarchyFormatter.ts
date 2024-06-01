class IssueHierarchyFormatter {

    static formatHierarchy(range: Range) {
        const issues: string[] = []

        const issuesRange = RangeUtils.toColumnRange(range, GSheetProjectSettings.projectsIssueColumnName)
        if (issuesRange != null) {
            issuesRange.getValues()
                .map(it => it[0]?.toString())
                .forEach(it => issues.push(it))
        }

        const parentIssuesRange = RangeUtils.toColumnRange(range, GSheetProjectSettings.projectsParentIssueColumnName)
        if (parentIssuesRange != null) {
            parentIssuesRange.getValues()
                .map(it => it[0]?.toString())
                .forEach(it => issues.push(it))
        }

        this.formatHierarchyForIssues(issues)
    }

    static formatHierarchyForAllIssues() {
        const issues: string[] = []

        const parentIssuesRange = SheetUtils.getColumnRange(
            GSheetProjectSettings.projectsSheetName,
            GSheetProjectSettings.projectsParentIssueColumnName,
            GSheetProjectSettings.firstDataRow,
        )
        parentIssuesRange.getValues()
            .map(it => it[0]?.toString())
            .forEach(it => issues.push(it))

        this.formatHierarchyForIssues(issues)
    }

    static formatHierarchyForIssues(issues: string[]) {
        issues
            .filter(it => it?.length)
            .filter(Utils.distinct)
            .forEach(issue => this.formatHierarchyForIssue(issue))
    }

    static formatHierarchyForIssue(issue: string) {
        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.projectsSheetName)
        ProtectionLocks.lockRowsWithProtection(sheet)

        const issuesRange = SheetUtils.getColumnRange(
            GSheetProjectSettings.projectsSheetName,
            GSheetProjectSettings.projectsIssueColumnName,
            GSheetProjectSettings.firstDataRow,
        )
        const parentIssuesRange = SheetUtils.getColumnRange(
            GSheetProjectSettings.projectsSheetName,
            GSheetProjectSettings.projectsParentIssueColumnName,
            GSheetProjectSettings.firstDataRow,
        )


        const issueRange = issuesRange.createTextFinder(issue)
            .ignoreDiacritics(false)
            .matchCase(true)
            .matchEntireCell(true)
            .findNext()
        if (issueRange == null) {
            return
        }

        let issueRow = issueRange.getRow()


        const issueTitleRange = sheet.getRange(
            issuesRange.getRow(),
            SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsTitleColumnName),
        )
        let indentLevel = Math.ceil(RangeUtils.getIndent(issueTitleRange) / GSheetProjectSettings.indent)

        const shouldIssueHaveIndent = sheet.getRange(
            issuesRange.getRow(),
            SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsParentIssueColumnName),
        ).getValue()?.toString()?.trim()?.length
        if (!shouldIssueHaveIndent && indentLevel > 0) {
            indentLevel = 0
            RangeUtils.setStringIndent(issueTitleRange, 0)
        }


        const childIssueRows = parentIssuesRange.createTextFinder(issue)
            .ignoreDiacritics(false)
            .matchCase(true)
            .matchEntireCell(true)
            .findAll()
            .map(it => it.getRow())
            .filter(it => it !== issueRow)
        if (!childIssueRows.length) {
            return
        }

        for (const row of childIssueRows) {
            const currentGroupDepth = sheet.getRowGroupDepth(row)
            const expectedGroupDepth = Math.min(indentLevel + 1, 4)
            if (currentGroupDepth !== expectedGroupDepth) {
                sheet.getRange(row, 1).shiftRowGroupDepth(expectedGroupDepth - currentGroupDepth)
            }
        }

        const childIssueRanges: Range[] = []
        for (let rowIndex = 0; rowIndex < childIssueRows.length; ++rowIndex) {
            const row = childIssueRows[rowIndex]
            let lastRow = row
            for (; rowIndex < childIssueRanges.length && childIssueRows[rowIndex] === lastRow + 1; ++rowIndex) {
                ++lastRow
            }
            const combinedRange = sheet.getRange(row, 1, lastRow - row + 1, 1)
            childIssueRanges.push(combinedRange)
        }

        for (const childIssueRange of childIssueRanges) {
            const childIssueTitleRange = sheet.getRange(
                childIssueRange.getRow(),
                SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsTitleColumnName),
                childIssueRange.getNumRows(),
                1,
            )
            RangeUtils.setStringIndent(childIssueTitleRange, (indentLevel + 1) * GSheetProjectSettings.indent)
        }


        // move children after the issue:
        let lastIssueOrConnectedChildIssueRow = issueRow
        for (const childIssueRange of childIssueRanges) {
            const childIssueRow = childIssueRange.getRow()
            if (childIssueRow === issueRow + 1) {
                lastIssueOrConnectedChildIssueRow = childIssueRow + childIssueRange.getNumRows()
                break
            }
        }

        for (const childIssueRange of childIssueRanges) {
            const childIssueRow = childIssueRange.getRow()
            if (childIssueRow < issueRow) {
                continue
            }

            if (childIssueRow < lastIssueOrConnectedChildIssueRow) {
                continue
            }

            sheet.moveRows(childIssueRange, lastIssueOrConnectedChildIssueRow + 1)
            lastIssueOrConnectedChildIssueRow += childIssueRange.getNumRows()
        }


        // move children before the issue:
        for (const childIssueRange of childIssueRanges.toReversed()) {
            const childIssueRow = childIssueRange.getRow()
            if (childIssueRow >= issueRow) {
                continue
            }

            sheet.moveRows(childIssueRange, issueRow + 1)
            issueRow -= childIssueRange.getNumRows()
        }
    }

}
