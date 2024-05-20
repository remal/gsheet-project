class SheetUtils {

    static findSheetByName(sheetName: string): Sheet | undefined {
        if (!sheetName?.length) {
            return undefined
        }

        sheetName = Utils.normalizeName(sheetName)
        return ExecutionCache.getOrComputeCache(['findSheetByName', sheetName], () => {
            for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
                const name = Utils.normalizeName(sheet.getSheetName())
                if (name === sheetName) {
                    return sheet
                }
            }
            return undefined
        })
    }

    static getSheetByName(sheetName: string): Sheet {
        return this.findSheetByName(sheetName) ?? (() => {
            throw new Error(`"${sheetName}" sheet can't be found`)
        })()
    }

    static findColumnByName(
        sheet: Sheet | string | null | undefined,
        columnName: string | null | undefined,
    ): number | undefined {
        if (!columnName?.length) {
            return undefined
        }

        if (Utils.isString(sheet)) {
            sheet = this.findSheetByName(sheet)
        }
        if (sheet == null || !this.isGridSheet(sheet)) {
            return undefined
        }

        columnName = Utils.normalizeName(columnName)
        return ExecutionCache.getOrComputeCache(['findColumnByName', sheet, columnName], () => {
            for (const col of Utils.range(1, sheet.getLastColumn())) {
                const name = Utils.normalizeName(sheet.getRange(1, col).getValue())
                if (name === columnName) {
                    return col
                }
            }

            return undefined
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

    static getColumnRange(sheet: Sheet | string, column: string | number, fromRow?: number): Range {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }
        if (Utils.isString(column)) {
            column = this.getColumnByName(sheet, column)
        }
        if (fromRow == null) {
            fromRow = 1
        }

        const lastRow = sheet.getLastRow()
        if (fromRow > lastRow) {
            return sheet.getRange(fromRow, column)
        }

        const rows = lastRow - fromRow + 1
        return sheet.getRange(fromRow, column, rows, 1)
    }

    static getRowRange(sheet: Sheet | string, row: number, fromColumn?: number | string): Range {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }
        if (fromColumn == null) {
            fromColumn = 1
        } else if (Utils.isString(fromColumn)) {
            fromColumn = this.getColumnByName(sheet, fromColumn)
        }

        const lastColumn = sheet.getLastColumn()
        if (fromColumn > lastColumn) {
            return sheet.getRange(row, fromColumn)
        }

        const columns = lastColumn - fromColumn + 1
        return sheet.getRange(row, fromColumn, 1, columns)
    }

    static isGridSheet(sheet: Sheet | string | null | undefined): boolean {
        if (Utils.isString(sheet)) {
            sheet = this.findSheetByName(sheet)
        }
        if (sheet == null) {
            return false
        }

        return sheet.getType()?.toString() === 'GRID'
    }

}
