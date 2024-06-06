class SheetUtils {

    static isGridSheet(sheet: Sheet | string | null | undefined): boolean {
        if (Utils.isString(sheet)) {
            sheet = this.findSheetByName(sheet)
        }
        if (sheet == null) {
            return false
        }

        return sheet.getType() === SpreadsheetApp.SheetType.GRID
    }

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

        ProtectionLocks.lockAllColumns(sheet)

        columnName = Utils.normalizeName(columnName)
        return ExecutionCache.getOrComputeCache(['findColumnByName', sheet, columnName], () => {
            for (const col of Utils.range(GSheetProjectSettings.titleRow, sheet.getLastColumn())) {
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

    static getColumnRange(sheet: Sheet | string, column: string | number, minRow?: number): Range {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }
        if (Utils.isString(column)) {
            column = this.getColumnByName(sheet, column)
        }
        if (minRow == null || minRow < 1) {
            minRow = 1
        }

        const lastRow = sheet.getLastRow()
        if (minRow > lastRow) {
            return sheet.getRange(minRow, column)
        }

        const rows = lastRow - minRow + 1
        return sheet.getRange(minRow, column, rows, 1)
    }

    static getColumnsValues<
        C extends Record<string, string | number>,
        R extends Record<keyof C, any[]>
    >(sheet: Sheet | string, columns: C, minRow?: number, maxRow?: number): R {
        const getter: (range: Range) => any[][] = range => range.getValues()
        return this._getColumnsProps(sheet, columns, getter, minRow, maxRow)
    }

    static getColumnsStringValues<
        C extends Record<string, string | number>,
        R extends Record<keyof C, string[]>
    >(sheet: Sheet | string, columns: C, minRow?: number, maxRow?: number): R {
        const getter: (range: Range) => any[][] = range => range.getValues()
        const result = this._getColumnsProps(sheet, columns, getter, minRow, maxRow)
        for (const [key, values] of Object.entries(result)) {
            (result as {})[key] = values.map(value => value.toString())
        }
        return result as R
    }

    static getColumnsFormulas<
        C extends Record<string, string | number>,
        R extends Record<keyof C, string[]>
    >(sheet: Sheet | string, columns: C, minRow?: number, maxRow?: number): R {
        const getter: (range: Range) => string[][] = range => range.getFormulas()
        return this._getColumnsProps(sheet, columns, getter, minRow, maxRow)
    }

    private static _getColumnsProps<
        C extends Record<string, string | number>,
        V extends any,
        R extends Record<keyof C, V[]>
    >(
        sheet: Sheet | string,
        columns: C,
        getter: (range: Range) => V[][],
        minRow: number | undefined,
        maxRow: number | undefined,
    ): R {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }
        if (!Object.keys(columns).length) {
            return {} as R
        }
        if (minRow == null || minRow < 1) {
            minRow = 1
        }
        if (maxRow == null) {
            maxRow = sheet.getLastRow()
        }

        const columnToNumber = Object.keys(columns)
            .reduce((rec, key) => {
                const value = columns[key]
                rec[key] = Utils.isString(value)
                    ? this.getColumnByName(sheet, value)
                    : value
                return rec
            }, {} as Record<string, number>)
        const numbers = Object.values(columnToNumber).filter(Utils.distinct()).toSorted(Utils.numericAsc())

        const result = {} as R
        Object.keys(columns).forEach(key => (result as {})[key] = [])

        while (numbers.length) {
            const baseColumn = numbers.shift()!
            let columnsCount = 1
            while (numbers.length) {
                const nextNumber = numbers[0]
                if (nextNumber === baseColumn + columnsCount) {
                    ++columnsCount
                    numbers.shift()
                } else {
                    break
                }
            }

            const range = sheet.getRange(
                minRow,
                baseColumn,
                Math.max(maxRow - minRow + 1, 1),
                columnsCount,
            )
            const props = getter(range)

            props.forEach(rows => rows.forEach((columnValue, index) => {
                const column = baseColumn + index
                for (const [columnKey, columnNumber] of Object.entries(columnToNumber)) {
                    if (column === columnNumber) {
                        result[columnKey].push(columnValue)
                    }
                }
            }))
        }
        return result
    }

    static getRowRange(sheet: Sheet | string, row: number, minColumn?: number | string): Range {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }
        if (minColumn == null) {
            minColumn = 1
        } else if (Utils.isString(minColumn)) {
            minColumn = this.getColumnByName(sheet, minColumn)
        } else if (minColumn < 1) {
            minColumn = 1
        }

        const lastColumn = sheet.getLastColumn()
        if (minColumn > lastColumn) {
            return sheet.getRange(row, minColumn)
        }

        const columns = lastColumn - minColumn + 1
        return sheet.getRange(row, minColumn, 1, columns)
    }

}
