class SheetUtils {

    static findSheetByName(sheetName: SheetName): Sheet | undefined {
        if (!sheetName?.length) {
            return undefined
        }

        const sheets = ExecutionCache.getOrCompute('sheets-by-name', () => {
            const result = new Map<SheetName, Sheet>()
            for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
                const name = Utils.normalizeName(sheet.getSheetName())
                result.set(name, sheet)
            }
            return result
        }, `${SheetUtils.name}: ${this.findSheetByName.name}`)

        sheetName = Utils.normalizeName(sheetName)
        return sheets.get(sheetName)
    }

    static getSheetByName(sheetName: string): Sheet {
        return this.findSheetByName(sheetName) ?? (() => {
            throw new Error(`"${sheetName}" sheet can't be found`)
        })()
    }

    static isGridSheet(sheet: Sheet | SheetName | null | undefined): boolean {
        if (Utils.isString(sheet)) {
            sheet = this.findSheetByName(sheet)
        }
        if (sheet == null) {
            return false
        }

        return sheet.getType() === SpreadsheetApp.SheetType.GRID
    }

    static getLastRow(sheet: Sheet | SheetName): Row {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }

        return ExecutionCache.getOrCompute(['last-row', sheet], () =>
            Math.max(sheet.getLastRow(), 1),
        )
    }

    static setLastRow(sheet: Sheet | SheetName, lastRow: Row) {
        ExecutionCache.put(['last-row', sheet], Math.max(lastRow, 1))
    }

    static getLastColumn(sheet: Sheet | SheetName): Column {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }

        return ExecutionCache.getOrCompute(['last-column', sheet], () =>
            Math.max(sheet.getLastColumn(), 1),
        )
    }

    static setLastColumn(sheet: Sheet | SheetName, lastColumn: Column) {
        ExecutionCache.put(['last-column', sheet], lastColumn)
    }

    static findColumnByName(
        sheet: Sheet | SheetName | null | undefined,
        columnName: ColumnName | null | undefined,
    ): Column | undefined {
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

        const columns = ExecutionCache.getOrCompute(['columns-by-name', sheet], () => {
            const result = new Map<ColumnName, Column>()
            for (const col of Utils.range(GSheetProjectSettings.titleRow, this.getLastColumn(sheet))) {
                const name = Utils.normalizeName(sheet.getRange(1, col).getValue())
                result.set(name, col)
            }
            return result
        }, `${SheetUtils.name}: ${this.findColumnByName.name}`)

        columnName = Utils.normalizeName(columnName)
        return columns.get(columnName)
    }

    static getColumnByName(sheet: Sheet | SheetName, columnName: ColumnName): Column {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }

        return this.findColumnByName(sheet, columnName) ?? (() => {
            throw new Error(`"${sheet.getSheetName()}" sheet: "${columnName}" column can't be found`)
        })()
    }

    static getColumnRange(sheet: Sheet | SheetName, column: ColumnName | Column, minRow?: Row): Range {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet)
        }
        if (Utils.isString(column)) {
            column = this.getColumnByName(sheet, column)
        }
        if (minRow == null || minRow < 1) {
            minRow = 1
        }

        const lastRow = this.getLastRow(sheet)
        if (minRow > lastRow) {
            return sheet.getRange(minRow, column)
        }

        const rows = lastRow - minRow + 1
        return sheet.getRange(minRow, column, rows, 1)
    }

    static getColumnsValues<
        C extends Record<string, ColumnName | Column>,
        R extends Record<keyof C, any[]>
    >(sheet: Sheet | SheetName, columns: C, minRow?: Row, maxRow?: Row): R {
        function getValues(range: Range): any[][] {
            return range.getValues()
        }

        return this._getColumnsProps(sheet, columns, getValues, minRow, maxRow)
    }

    static getColumnsStringValues<
        C extends Record<string, ColumnName | Column>,
        R extends Record<keyof C, string[]>
    >(sheet: Sheet | SheetName, columns: C, minRow?: Row, maxRow?: Row): R {
        function getValues(range: Range): any[][] {
            return range.getValues()
        }

        const result = this._getColumnsProps(sheet, columns, getValues, minRow, maxRow)
        for (const [key, values] of Object.entries(result)) {
            (result as {})[key] = values.map(value => value.toString())
        }
        return result as R
    }

    static getColumnsFormulas<
        C extends Record<string, ColumnName | Column>,
        R extends Record<keyof C, Formula[]>
    >(sheet: Sheet | SheetName, columns: C, minRow?: Row, maxRow?: Row): R {
        function getFormulas(range: Range): Formula[][] {
            return range.getFormulas()
        }

        return this._getColumnsProps(sheet, columns, getFormulas, minRow, maxRow)
    }

    private static _getColumnsProps<
        C extends Record<string, ColumnName | Column>,
        V extends any,
        R extends Record<keyof C, V[]>
    >(
        sheet: Sheet | SheetName,
        columns: C,
        getter: (range: Range) => V[][],
        minRow: Row | undefined,
        maxRow: Row | undefined,
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
            maxRow = this.getLastRow(sheet)
        }

        const columnToNumber = Object.keys(columns)
            .reduce((rec, key) => {
                const value = columns[key]
                rec[key] = Utils.isString(value)
                    ? this.getColumnByName(sheet, value)
                    : value
                return rec
            }, {} as Record<string, Column>)
        const numbers = Object.values(columnToNumber).filter(Utils.distinct()).toSorted(Utils.numericAsc())

        const result = {} as R
        Object.keys(columns).forEach(key => (result as {})[key] = [])
        if (minRow > maxRow) {
            return result
        }

        Observability.timed(
            [
                SheetUtils.name,
                this._getColumnsProps.name,
                sheet.getSheetName(),
                `rows from #${minRow} to #${maxRow}`,
                `columns #${numbers.join(', #')} (${getter.name})`,
            ].join(': '),
            () => {
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
            },
        )
        return result
    }

    static getRowRange(sheet: Sheet | SheetName, row: Row, minColumn?: ColumnName | Column): Range {
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

        const lastColumn = this.getLastColumn(sheet)
        if (minColumn > lastColumn) {
            return sheet.getRange(row, minColumn)
        }

        const columns = lastColumn - minColumn + 1
        return sheet.getRange(row, minColumn, 1, columns)
    }

}
