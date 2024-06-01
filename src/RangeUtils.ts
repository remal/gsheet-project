class RangeUtils {

    static toColumnRange(
        range: Range | null | undefined,
        column: number | string | null | undefined,
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

        if (!this.doesRangeHaveColumn(range, column)) {
            return undefined
        }

        return range.offset(
            0,
            column - range.getColumn(),
            range.getNumRows(),
            1,
        )
    }

    static doesRangeHaveColumn(
        range: Range | null | undefined,
        column: number | string | null | undefined,
    ): boolean {
        if (range == null || column == null) {
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

    static doesRangeIntersectsWithNamedRange(
        range: Range | null | undefined,
        namedRange: NamedRange | string | null | undefined,
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
