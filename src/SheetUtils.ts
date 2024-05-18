class SheetUtils {

    static findSheetByName(sheetName: string): Sheet | null {
        if (!sheetName?.length) {
            return null
        }

        sheetName = Utils.normalizeName(sheetName)
        return ExecutionCache.getOrComputeCache(['findSheetByName', sheetName], () => {
            for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
                const name = Utils.normalizeName(sheet.getSheetName())
                if (name === sheetName) {
                    return sheet
                }
            }
            return null
        })
    }

    static getSheetByName(sheetName: string): Sheet {
        return this.findSheetByName(sheetName) ?? (() => {
            throw new Error(`"${sheetName}" sheet can't be found`)
        })()
    }

    static findColumnByName(sheet: Sheet | string | null, columnName: string | null): number | null {
        if (!columnName?.length) {
            return null
        }

        if (Utils.isString(sheet)) {
            sheet = this.findSheetByName(sheet)
        }
        if (sheet == null) {
            return null
        }

        columnName = Utils.normalizeName(columnName)
        return ExecutionCache.getOrComputeCache(['findColumnByName', sheet, columnName], () => {
            for (const col of Utils.range(1, sheet.getLastColumn())) {
                const name = Utils.normalizeName(sheet.getRange(1, col).getValue())
                if (name === columnName) {
                    return col
                }
            }

            return null
        })
    }

    static getColumnByName(sheet: Sheet | string, columnName: string): number {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }

        return this.findColumnByName(sheet, columnName) ?? (() => {
            throw new Error(`"${sheet.getSheetName()}" sheet: "${columnName}" column can't be found`)
        })()
    }

}
