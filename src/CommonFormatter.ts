class CommonFormatter {

    static setMiddleVerticalAlign() {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
                sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
                    .setVerticalAlignment('middle')
            })
    }

    static onChange(event?: SheetsOnChange) {
        if (['INSERT_ROW', 'INSERT_COLUMN'].includes(event?.changeType?.toString() ?? '')) {
            this.setMiddleVerticalAlign()
        }
    }

}
