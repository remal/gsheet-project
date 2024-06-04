abstract class SheetLayout {

    protected abstract get sheetName(): string

    protected abstract get columns(): ReadonlyArray<ColumnInfo>

    protected get sheet(): Sheet {
        const sheetName = this.sheetName
        let sheet = SheetUtils.findSheetByName(sheetName)
        if (sheet == null) {
            sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName)
            ExecutionCache.resetCache()
        }
        return sheet
    }

    private get _documentFlagPrefix(): string {
        return `${this.constructor?.name || Utils.normalizeName(this.sheetName)}:migrateColumns:`
    }

    private get _documentFlag(): string {
        return `${this._documentFlagPrefix}$$$HASH$$$:${GSheetProjectSettings.computeStringSettingsHash()}`
    }

    migrateColumnsIfNeeded() {
        if (DocumentFlags.isSet(this._documentFlag)) {
            return
        }

        this.migrateColumns()
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


        const sheet = this.sheet
        ProtectionLocks.lockColumnsWithProtection(sheet)

        let lastColumn = Math.max(sheet.getLastColumn(), 1)
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

                if (info.hiddenByDefault) {
                    sheet.hideColumns(lastColumn)
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

            const range = sheet.getRange(
                GSheetProjectSettings.firstDataRow,
                column,
                maxRows,
                1,
            )
            if (info.rangeName?.length) {
                SpreadsheetApp.getActiveSpreadsheet().setNamedRange(info.rangeName, range)
            }

            let dataValidation: (DataValidation | null) = info.dataValidation?.call(info) ?? null
            if (dataValidation != null) {
                if (dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CUSTOM_FORMULA) {
                    const formula = dataValidation.getCriteriaValues()[0].toString()
                        .replaceAll(/#SELF\b/g, 'INDIRECT(ADDRESS(ROW(), COLUMN()))')
                        .split(/[\r\n]+/)
                        .map(line => line.trim())
                        .filter(line => line.length)
                        .map(line => line + (line.endsWith(',') || line.endsWith(';') ? ' ' : ''))
                        .join('')
                    dataValidation = dataValidation.copy()
                        .requireFormulaSatisfied(formula)
                        .build()
                }
            }
            range.setDataValidation(dataValidation)
        }

        DocumentFlags.set(this._documentFlag)
        DocumentFlags.cleanupByPrefix(this._documentFlagPrefix)

        const waitForAllDataExecutionsCompletion = SpreadsheetApp.getActiveSpreadsheet()['waitForAllDataExecutionsCompletion']
        if (Utils.isFunction(waitForAllDataExecutionsCompletion)) {
            try {
                waitForAllDataExecutionsCompletion(10)
            } catch (e) {
                console.warn(e)
            }
        }
    }

}

interface ColumnInfo {
    name: string
    arrayFormula?: string
    rangeName?: string
    dataValidation?: () => (DataValidation | null)
    defaultFontSize?: number
    defaultWidth?: number | WidthString
    hiddenByDefault?: boolean
}

type WidthString = '#height' | '#default-height'
