class RangeUtils {

    static isRangeSheet(
        range: Range | null | undefined,
        sheet: Sheet | SheetName | null | undefined,
    ): boolean {
        if (range == null) {
            return false
        }

        if (Utils.isString(sheet)) {
            sheet = SheetUtils.findSheetByName(sheet)
        }
        if (sheet == null) {
            return false
        }

        return range.getSheet().getSheetId() === sheet.getSheetId()
    }

    static getAbsoluteA1Notation(range: Range): string {
        return range.getA1Notation()
            .replaceAll(/[A-Z]+/g, '$$$&')
            .replaceAll(/\d+/g, '$$$&')
    }

    static getAbsoluteReferenceFormula(range: Range): string {
        return '=' + this.getAbsoluteA1Notation(range)
    }

    static toColumnRange(
        range: Range | null | undefined,
        column: ColumnName | Column | null | undefined,
    ): Range | undefined {
        if (range == null || column == null) {
            return undefined
        }

        if (Utils.isString(column)) {
            column = SheetUtils.findColumnByName(range.getSheet(), column)
        }
        if (column == null) {
            return undefined
        }

        return range.offset(
            0,
            column - range.getColumn(),
            range.getNumRows(),
            1,
        )
    }

    static withMinRow(range: Range, minRow: Row): Range {
        const startRow = range.getRow()
        const rowDiff = minRow - startRow
        if (rowDiff <= 0) {
            return range
        }

        return range.offset(
            rowDiff,
            0,
            Math.max(range.getNumRows() - rowDiff, 1),
            range.getNumColumns(),
        )
    }

    static withMaxRow(range: Range, maxRow: Row): Range {
        const startRow = range.getRow()
        const endRow = startRow + range.getNumRows() - 1
        if (maxRow >= endRow) {
            return range
        }

        return range.offset(
            0,
            0,
            Math.max(maxRow - startRow + 1, 1),
            range.getNumColumns(),
        )
    }

    static withMinMaxRows(range: Range, minRow: Row, maxRow: Row): Range {
        range = this.withMinRow(range, minRow)
        range = this.withMaxRow(range, maxRow)
        return range
    }

    static doesRangeHaveColumn(
        range: Range | null | undefined,
        column: ColumnName | Column | null | undefined,
    ): boolean {
        if (range == null) {
            return false
        }

        if (Utils.isString(column)) {
            column = SheetUtils.findColumnByName(range.getSheet(), column)
        }
        if (column == null) {
            return false
        }

        const minColumn = range.getColumn()
        const maxColumn = minColumn + range.getNumColumns() - 1
        return minColumn <= column && column <= maxColumn
    }

    static doesRangeHaveSheetColumn(
        range: Range | null | undefined,
        sheet: Sheet | SheetName | null | undefined,
        column: ColumnName | Column | null | undefined,
    ): boolean {
        return this.isRangeSheet(range, sheet) && this.doesRangeHaveColumn(range, column)
    }

    static doesRangeIntersectsWithNamedRange(
        range: Range | null | undefined,
        namedRange: NamedRange | RangeName | null | undefined,
    ): boolean {
        if (range == null || namedRange == null) {
            return false
        }

        if (Utils.isString(namedRange)) {
            namedRange = NamedRangeUtils.findNamedRange(namedRange)
        }
        if (namedRange == null) {
            return false
        }

        const rangeToFind = namedRange.getRange()
        if (range.getSheet().getSheetId() !== namedRange.getRange().getSheet().getSheetId()) {
            return false
        }

        const minColumnToFind = rangeToFind.getColumn()
        const maxColumnToFind = minColumnToFind + rangeToFind.getNumColumns() - 1
        const minColumn = range.getColumn()
        const maxColumn = minColumn + range.getNumColumns() - 1
        if (maxColumnToFind < minColumn || maxColumn < minColumnToFind) {
            return false
        }

        const minRowToFind = rangeToFind.getRow()
        const maxRowToFind = minRowToFind + rangeToFind.getNumRows() - 1
        const minRow = range.getRow()
        const maxRow = minRow + range.getNumRows() - 1
        if (maxRowToFind < minRow || maxRow < minRowToFind) {
            return false
        }

        return true
    }

    static getIndent(range: Range): number {
        const numberFormat = range.getNumberFormat()
        return this._parseIndent(numberFormat)
    }

    static setIndent(range: Range, indent: number) {
        indent = Math.max(indent, 0)

        let numberFormat = range.getNumberFormat()
        if (indent === this._parseIndent(numberFormat)) {
            return
        }

        numberFormat = numberFormat.trim()
        if (numberFormat.length) {
            range.setNumberFormat(`${' '.repeat(indent)}${numberFormat}`)
        } else if (indent > 0) {
            range.setNumberFormat(`${' '.repeat(indent)}@`)
        } else {
            // do nothing
        }
    }

    static setStringIndent(range: Range, indent: number) {
        indent = Math.max(indent, 0)
        range.setNumberFormat(`${' '.repeat(indent)}@`)
    }

    private static _parseIndent(numberFormat: string): number {
        const indentMatch = numberFormat.match(/^( +)/)
        if (indentMatch) {
            return indentMatch[0].length
        }
        return 0
    }

}
