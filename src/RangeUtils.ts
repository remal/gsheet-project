class RangeUtils {

    static doesRangeHaveColumn(range: Range | null, columnName: string): boolean {
        if (range == null) {
            return false
        }

        const sheet = range.getSheet()
        const columnToFind = SheetUtils.findColumnByName(sheet, columnName)
        if (columnToFind == null) {
            return false
        }

        for (const y of Utils.range(1, range.getHeight())) {
            let hasMerge = false
            for (const x of Utils.range(1, range.getWidth())) {
                const cell = range.getCell(y, x)
                if (cell.isPartOfMerge()) {
                    hasMerge = true
                }

                const col = cell.getColumn()
                if (col === columnToFind) {
                    return true
                }
            }

            if (!hasMerge) {
                break
            }
        }

        return false
    }

}
