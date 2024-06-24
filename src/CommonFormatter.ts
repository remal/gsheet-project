class CommonFormatter {

    static applyCommonFormatsToAllSheets() {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
                this.highlightCellsWithFormula(sheet)

                const range = SheetUtils.getWholeSheetRange(sheet)
                this.applyCommonFormatsToRange(range)
            })
    }

    static highlightCellsWithFormula(sheet: Sheet | SheetName) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet)
        }

        const range = SheetUtils.getWholeSheetRange(sheet)
        ConditionalFormatting.addConditionalFormatRule(
            range,
            {
                scope: 'common',
                order: 10_000,
                configurer: builder => builder
                    .whenFormulaSatisfied(`
                        =ISFORMULA(A1)
                    `)
                    .setItalic(true)
                    .setFontColor('#333'),
            },
        )
    }

    static applyCommonFormatsToRowRange(range: Range) {
        const sheet = range.getSheet()
        range = sheet.getRange(
            range.getRow(),
            1,
            range.getNumRows(),
            SheetUtils.getMaxColumns(sheet),
        )
        this.applyCommonFormatsToRange(range)
    }

    static applyCommonFormatsToRange(range: Range) {
        range
            .setVerticalAlignment('middle')
    }

}
