abstract class SheetLayout {

    protected abstract get sheetName(): string

    protected abstract get columns(): ReadonlyArray<ColumnInfo>

    protected get sheet(): Sheet {
        return SheetUtils.getSheetByName(this.sheetName)
    }

    migrateColumns() {
        const columns = this.columns.reduce(
            (map, info) => {
                map.set(Utils.normalizeName(info.name), info)
                return map
            },
            new Map<string, ColumnInfo>(),
        )
        if (!columns.size) {
            return
        }

        const documentFlagPrefix = `${this.constructor?.name || Utils.normalizeName(this.sheetName)}:migrateColumns:`
        const documentFlag = `${documentFlagPrefix}$$$HASH$$$:${GSheetProjectSettings.computeStringSettingsHash()}`
        if (DocumentFlags.isSet(documentFlag)) {
            return
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
                const titleRange = sheet.getRange(GSheetProjectSettings.titleRow, lastColumn)
                    .setValue(info.name)

                if (info.defaultFontSize) {
                    titleRange.setFontSize(info.defaultFontSize)
                }

                if (Utils.isNumber(info.defaultWidth)) {
                    sheet.setColumnWidth(lastColumn, info.defaultWidth)
                } else if (info.defaultWidth === '#default-height') {
                    sheet.setColumnWidth(lastColumn, 21)
                } else if (info.defaultWidth === '#height') {
                    const height = sheet.getRowHeight(1)
                    sheet.setColumnWidth(lastColumn, height)
                }

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
                    .map(line => line + (line.endsWith(',') || line.endsWith(';') ? ' ' : ''))
                    .join('')
                const formulaToExpect = `={"${Utils.escapeFormulaString(info.name)}"; ${arrayFormulaNormalized}}`
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
                    maxRows,
                    1,
                ))
            }
        }

        DocumentFlags.set(documentFlag)
        DocumentFlags.cleanupByPrefix(documentFlagPrefix)

        const waitForAllDataExecutionsCompletion = SpreadsheetApp.getActiveSpreadsheet()['waitForAllDataExecutionsCompletion']
        if (Utils.isFunction(waitForAllDataExecutionsCompletion)) {
            waitForAllDataExecutionsCompletion(10)
        }
    }

}

interface ColumnInfo {
    name: string
    arrayFormula?: string
    rangeName?: string
    defaultFontSize?: number
    defaultWidth?: number | WidthString
}

type WidthString = '#height' | '#default-height'
