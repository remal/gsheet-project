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
        issues = issues
            .filter(it => it?.length)
            .filter(Utils.distinct())
        issues.forEach(issue => this.formatHierarchyForIssue(issue))
    }

    static formatHierarchyForIssue(issue: string) {
        const issueSlug = issue.replaceAll(/[\r\n]+/g, '').replace(/^(.{0,25}).*$/, '$1')
        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.projectsSheetName)
        ProtectionLocks.lockRowsWithProtection(sheet)


        const issueRange = SheetUtils.getColumnRange(
            GSheetProjectSettings.projectsSheetName,
            GSheetProjectSettings.projectsIssueColumnName,
            GSheetProjectSettings.firstDataRow,
        )
            .createTextFinder(issue)
            .ignoreDiacritics(false)
            .matchCase(true)
            .matchEntireCell(true)
            .findNext()
        if (issueRange == null) {
            return
        }

        let issueRow = issueRange.getRow()


        const issueTitleRange = sheet.getRange(
            issueRow,
            SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsTitleColumnName),
        )
        let indentLevel = Math.ceil(RangeUtils.getIndent(issueTitleRange) / GSheetProjectSettings.indent)

        const shouldIssueHaveIndent = sheet.getRange(
            issueRow,
            SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsParentIssueColumnName),
        ).getValue()?.toString()?.trim()?.length
        if (!shouldIssueHaveIndent && indentLevel > 0) {
            indentLevel = 0
            RangeUtils.setStringIndent(issueTitleRange, 0)
        }


        const childIssueRows = SheetUtils.getColumnRange(
            GSheetProjectSettings.projectsSheetName,
            GSheetProjectSettings.projectsParentIssueColumnName,
            GSheetProjectSettings.firstDataRow,
        )
            .createTextFinder(issue)
            .ignoreDiacritics(false)
            .matchCase(true)
            .matchEntireCell(true)
            .findAll()
            .map(it => it.getRow())
            .filter(it => it !== issueRow)
        if (!childIssueRows.length) {
            return
        }

        Utils.timed(`${IssueHierarchyFormatter.name}: ${issueSlug}: Adjust groups`, () => {
            for (const row of childIssueRows) {
                const currentGroupDepth = sheet.getRowGroupDepth(row)
                const expectedGroupDepth = Math.min(indentLevel + 1, 4)
                if (currentGroupDepth !== expectedGroupDepth) {
                    sheet.getRange(row, 1).shiftRowGroupDepth(expectedGroupDepth - currentGroupDepth)
                }
            }
        })

        const childIssueRanges: Range[] = []
        for (let rowIndex = 0; rowIndex < childIssueRows.length; ++rowIndex) {
            const row = childIssueRows[rowIndex]

            let rows = 1
            let lastRow = row
            for (let i = rowIndex + 1; i < childIssueRows.length; ++i) {
                const nextRow = childIssueRows[i]
                if (nextRow === lastRow + 1) {
                    ++rows
                    lastRow = nextRow
                } else {
                    break
                }
            }
            rowIndex += rows - 1

            const combinedRange = sheet.getRange(row, 1, rows, 1)
            childIssueRanges.push(combinedRange)
        }

        Utils.timed(`${IssueHierarchyFormatter.name}: ${issueSlug}: Adjust indents`, () => {
            for (const childIssueRange of childIssueRanges) {
                const childIssueTitleRange = sheet.getRange(
                    childIssueRange.getRow(),
                    SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsTitleColumnName),
                    childIssueRange.getNumRows(),
                    1,
                )
                RangeUtils.setStringIndent(childIssueTitleRange, (indentLevel + 1) * GSheetProjectSettings.indent)
            }
        })


        if (childIssueRanges.length === 1
            && childIssueRanges[0].getRow() === issueRow + 1
        ) {
            return
        }


        // move children after the issue:
        Utils.timed(`${IssueHierarchyFormatter.name}: ${issueSlug}: Move children after the issue`, () => {
            let lastIssueOrConnectedChildIssueRow = issueRow
            for (const childIssueRange of childIssueRanges) {
                const childIssueRow = childIssueRange.getRow()
                if (childIssueRow === issueRow + 1) {
                    lastIssueOrConnectedChildIssueRow += childIssueRange.getNumRows()
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
        })


        // move children before the issue:
        Utils.timed(`${IssueHierarchyFormatter.name}: ${issueSlug}: Move children before the issue`, () => {
            for (const childIssueRange of childIssueRanges.toReversed()) {
                const childIssueRow = childIssueRange.getRow()
                if (childIssueRow >= issueRow) {
                    continue
                }

                sheet.moveRows(childIssueRange, issueRow + 1)
                issueRow -= childIssueRange.getNumRows()
            }
        })
    }

}
