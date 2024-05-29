abstract class AbstractSheetLayout {

    protected abstract get sheetName(): string

    protected abstract get columns(): ColumnInfo[]

    protected get sheet(): Sheet {
        return SheetUtils.getSheetByName(this.sheetName)
    }

    migrateColumns() {
        const columns = this.columns.reduce(
            (map, info) => map.set(Utils.normalizeName(info.name), info),
            new Map<string, ColumnInfo>(),
        )
        if (!columns.size) {
            return
        }

        const cacheKey = `SheetLayout:migrateColumns:$$$HASH$$$:${GSheetProjectSettings.computeSettingsHash()}:${this.sheetName}`
        const cache = CacheService.getDocumentCache()
        if (cache != null) {
            if (cache.get(cacheKey) === 'true') {
                return
            }
        }

        const sheet = this.sheet
        ProtectionLocks.lockColumnsWithProtection(sheet)

        let lastColumn = sheet.getLastColumn()
        const maxRows = sheet.getMaxRows()
        const existingNormalizedNames = sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn)
            .getValues()[0]
            .map(it => it?.toString())
            .map(it => it?.length ? Utils.normalizeName(it) : '')
        for (const [columnName, info] of columns.entries()) {
            if (!existingNormalizedNames.includes(columnName)) {
                sheet.getRange(GSheetProjectSettings.titleRow, lastColumn)
                    .setValue(info.name)

                existingNormalizedNames.push(columnName)

                ++lastColumn
            }
        }

        const existingFormulas = new Lazy(() =>
            sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn).getFormulas()[0],
        )
        for (const [columnName, info] of columns.entries()) {
            const index = existingNormalizedNames.indexOf(columnName)
            if (index < 0) {
                continue
            }

            const column = index + 1

            if (info.arrayFormula?.length) {
                const arrayFormulaNormalized = info.arrayFormula.split(/[\r\n]+/)
                    .map(line => line.trim())
                    .filter(line => line.length)
                    .join('')
                    .trim()
                const formulaToExpect = `={"${Utils.escapeFormulaString(info.name)}", ${arrayFormulaNormalized}`
                const formula = existingFormulas.get()[index]
                if (formula !== formulaToExpect) {
                    sheet.getRange(GSheetProjectSettings.titleRow, column)
                        .setFormula(formulaToExpect)
                }
            }

            if (info.rangeName?.length) {
                SpreadsheetApp.getActiveSpreadsheet().setNamedRange(info.rangeName, sheet.getRange(
                    GSheetProjectSettings.firstDataRow,
                    column,
                    maxRows - GSheetProjectSettings.firstDataRow,
                    1,
                ))
            }
        }

        if (cache != null) {
            cache.put(cacheKey, 'true')
        }
    }

}

interface ColumnInfo {
    name: string
    arrayFormula?: string
    rangeName?: string
}
