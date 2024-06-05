class DefaultFormulas {

    static insertDefaultFormulas(range: Range) {
        if (!RangeUtils.doesRangeHaveSheetColumn(
                range,
                GSheetProjectSettings.sheetName,
                GSheetProjectSettings.issueColumnName,
            )
            && !RangeUtils.doesRangeHaveSheetColumn(
                range,
                GSheetProjectSettings.sheetName,
                GSheetProjectSettings.titleColumnName,
            )
        ) {
            return
        }

        const startRow = range.getRow()
        const endRow = startRow + range.getNumRows()
        ProtectionLocks.lockRows(range.getSheet(), endRow)
        for (let row = startRow; row <= endRow; ++row) {

        }
    }

}
