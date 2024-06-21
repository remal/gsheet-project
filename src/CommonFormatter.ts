class CommonFormatter {

    static applyCommonFormatsToAllSheets() {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
                this.highlightCellsWithFormula(sheet)

                const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
                this.applyCommonFormatsToRange(range)
            })
    }

    static highlightCellsWithFormula(sheet: Sheet | SheetName) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet)
        }

        ConditionalFormatting.addConditionalFormatRule(
            sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()),
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
            sheet.getMaxColumns(),
        )
        this.applyCommonFormatsToRange(range)
    }

    static applyCommonFormatsToRange(range: Range) {
        range
            .setVerticalAlignment('middle')
    }

}
