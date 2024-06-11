class CommonFormatter {

    static applyCommonFormatsToAllSheets() {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
                this.setMiddleVerticalAlign(sheet)
                this.setClipWrapStrategy(sheet)
                this.highlightCellsWithFormula(sheet)
            })
    }

    static setMiddleVerticalAlign(sheet: Sheet | string) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet)
        }

        sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
            .setVerticalAlignment('middle')
    }

    static setClipWrapStrategy(sheet: Sheet | string) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet)
        }

        sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    }

    static highlightCellsWithFormula(sheet: Sheet | string) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet)
        }

        ConditionalFormatting.addConditionalFormatRule(
            sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()),
            {
                scope: 'common',
                order: 10_000,
                configurer: builder => builder
                    .whenFormulaSatisfied('=ISFORMULA(A1)')
                    .setItalic(true)
                    .setFontColor('#333'),
            },
        )
    }

}
