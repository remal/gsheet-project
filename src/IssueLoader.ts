class IssueLoader {

    static loadIssues(range: Range) {
        if (!RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.issueIdColumnName)) {
            return
        }

        const sheet = range.getSheet()
        const rows = Array.from(Utils.range(1, range.getHeight()))
            .map(y => range.getCell(y, 1).getRow())
            .filter(row => row >= GSheetProjectSettings.firstDataRow)
            .filter(Utils.distinct)
        for (const row of rows) {
            this.loadIssuesForRow(sheet, row)
        }
    }

    static loadAllIssues() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            const hasIssueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.issueIdColumnName) != null
            if (!hasIssueIdColumn) {
                continue
            }

            for (const row of Utils.range(GSheetProjectSettings.firstDataRow, sheet.getLastRow())) {
                this.loadIssuesForRow(sheet, row)
            }
        }
    }

    private static loadIssuesForRow(sheet: Sheet, row: number) {
        if (row < GSheetProjectSettings.firstDataRow) {
            return
        }

        const issueIdColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueIdColumnName)
        const issueIdRange = sheet.getRange(row, issueIdColumn)
        const issueIds = GSheetProjectSettings.issueIdsExtractor(issueIdRange.getValue())
        if (!issueIds?.length) {
            return
        }

        console.log(`"${sheet.getSheetName()}" sheet: processing row #${row}`)
        issueIdRange.setBackground('#eee')
        try {
            const rootIssues = GSheetProjectSettings.issuesLoader(issueIds)
            const childIssues = new Lazy(() => {
                return GSheetProjectSettings.childIssuesLoader(issueIds)
                    .filter(issue => !issueIds.includes(GSheetProjectSettings.issueIdGetter(issue)))
            })
            const blockerIssues = new Lazy(() => {
                const ids = rootIssues.concat(childIssues.get())
                    .map(issue => GSheetProjectSettings.issueIdGetter(issue))
                return GSheetProjectSettings.blockerIssuesLoader(ids)
                    .filter(issue => !issueIds.includes(GSheetProjectSettings.issueIdGetter(issue)))
            })

            const isDoneColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.isDoneColumnName)
            if (isDoneColumn != null) {
                const isDone = GSheetProjectSettings.idDoneCalculator(rootIssues, childIssues.get())
                sheet.getRange(row, isDoneColumn).setValue(isDone ? 'Yes' : '')
            }

            for (const [columnName, getter] of Object.entries(GSheetProjectSettings.stringFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName)
                if (fieldColumn != null) {
                    sheet.getRange(row, fieldColumn).setValue(rootIssues
                        .map(getter)
                        .join('\n'),
                    )
                }
            }

            for (const [columnName, getter] of Object.entries(GSheetProjectSettings.booleanFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName)
                if (fieldColumn != null) {
                    const isTrue = rootIssues.every(getter)
                    sheet.getRange(row, fieldColumn).setValue(isTrue ? 'Yes' : '')
                }
            }

            const calculateIssueMetrics = (metricsIssues: Lazy<Issue[]>, metrics: IssueMetric[]) => {
                for (const metric of metrics) {
                    const metricColumn = SheetUtils.findColumnByName(sheet, metric.columnName)
                    if (metricColumn == null) {
                        continue
                    }

                    const metricRange = sheet.getRange(row, metricColumn)

                    const foundIssues = metricsIssues.get().filter(metric.filter)
                    if (!foundIssues.length) {
                        metricRange.clearContent().setFontColor(null)
                        continue
                    }

                    const metricIssueIds = foundIssues.map(issue => GSheetProjectSettings.issueIdGetter(issue))
                    const link = GSheetProjectSettings.issueIdsToUrl?.call(null, metricIssueIds)
                    if (link != null) {
                        metricRange.setFormula(`=HYPERLINK("${link}", "${foundIssues.length}")`)
                    } else {
                        metricRange.setFormula(`="${foundIssues.length}"`)
                    }

                    metricRange.setFontColor(metric.color)
                }
            }

            calculateIssueMetrics(childIssues, GSheetProjectSettings.childIssueMetrics)
            calculateIssueMetrics(blockerIssues, GSheetProjectSettings.blockerIssueMetrics)

        } finally {
            issueIdRange.setBackground(null)
        }
    }

}
