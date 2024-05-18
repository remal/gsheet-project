class IssueInfoLoader {

    private settings: GSheetProjectSettings

    constructor(settings: GSheetProjectSettings) {
        this.settings = settings;
    }


    loadIssueInfo(range: Range) {
        if (!RangeUtils.doesRangeHaveColumn(range, this.settings.issueIdColumnName)) {
            return
        }

        const sheet = range.getSheet()
        const rows = Array.from(Utils.range(1, range.getHeight()))
            .map(y => range.getCell(y, 1).getRow())
            .filter(row => row >= DATA_FIRST_ROW)
            .filter(Utils.distinct)
        for (const row of rows) {
            this.loadIssueInfoForRow(sheet, row)
        }
    }

    loadAllIssueInfo() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            const hasIssueIdColumn = SheetUtils.findColumnByName(sheet, this.settings.issueIdColumnName) != null
            if (!hasIssueIdColumn) {
                return
            }

            for (const row of Utils.range(DATA_FIRST_ROW, sheet.getLastRow())) {
                this.loadIssueInfoForRow(sheet, row)
            }
        }
    }

    private loadIssueInfoForRow(sheet: Sheet, row: number) {
        if (row < DATA_FIRST_ROW
            || sheet.isRowHiddenByUser(row)
        ) {
            return
        }

        const issueIdColumn = SheetUtils.getColumnByName(sheet, this.settings.issueIdColumnName)
        const issueIdRange = sheet.getRange(row, issueIdColumn)
        const issueIds = this.settings.issueIdsExtractor(issueIdRange.getValue())
        if (!issueIds.length) {
            return
        }

        console.log(`"${sheet.getSheetName()}" sheet: processing row #${row}`)
        issueIdRange.setBackground('#eee')
        try {
            const rootIssues = this.settings.issuesLoader(issueIds)
            const childIssues = new Lazy(() =>
                this.settings.childIssuesLoader(issueIds)
                    .filter(issue => !issueIds.includes(this.settings.issueIdGetter(issue))),
            )
            const blockerIssues = new Lazy(() =>
                this.settings.blockerIssuesLoader(
                    rootIssues.concat(childIssues.get())
                        .map(issue => this.settings.issueIdGetter(issue)),
                ),
            )

            const isDoneColumn = SheetUtils.findColumnByName(sheet, this.settings.isDoneColumnName)
            if (isDoneColumn != null) {
                const isDone = this.settings.idDoneCalculator(rootIssues, childIssues.get())
                sheet.getRange(row, isDoneColumn).setValue(isDone ? 'Yes' : '')
            }

            for (const [columnName, getter] of Object.entries(this.settings.stringFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName)
                if (fieldColumn != null) {
                    sheet.getRange(row, fieldColumn).setValue(rootIssues
                        .map(getter)
                        .join('\n'),
                    )
                }
            }

            for (const [columnName, getter] of Object.entries(this.settings.booleanFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName)
                if (fieldColumn != null) {
                    const isTrue = rootIssues.some(getter)
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

                    const metricIssueIds = foundIssues.map(issue => this.settings.issueIdGetter(issue))
                    const link = this.settings.issueIdsToUrl?.call(null, metricIssueIds)
                    if (link != null) {
                        metricRange.setFormula(`=HYPERLINK("${link}", "${foundIssues.length}")`)
                    } else {
                        metricRange.setFormula(`="${foundIssues.length}"`)
                    }

                    if (metric.color != null) {
                        metricRange.setFontColor(metric.color)
                    }
                }
            }

            calculateIssueMetrics(childIssues, this.settings.childIssueMetrics)
            calculateIssueMetrics(blockerIssues, this.settings.blockerIssueMetrics)

        } finally {
            issueIdRange.setBackground(null)
        }
    }

}
