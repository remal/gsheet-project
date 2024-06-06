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
        return `${this.constructor?.name || Utils.normalizeName(this.sheetName)}:migrate:`
    }

    private get _documentFlag(): string {
        return `${this._documentFlagPrefix}$$$HASH$$$:${GSheetProjectSettings.computeStringSettingsHash()}`
    }

    migrateIfNeeded() {
        if (DocumentFlags.isSet(this._documentFlag)) {
            return
        }

        this.migrate()
    }

    migrate() {
        const sheet = this.sheet
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


        ProtectionLocks.lockAllColumns(sheet)

        let lastColumn = sheet.getLastColumn()
        const maxRows = sheet.getMaxRows()
        const existingNormalizedNames = sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn)
            .getValues()[0]
            .map(it => it?.toString())
            .map(it => it?.length ? Utils.normalizeName(it) : '')
        for (const [columnName, info] of columns.entries()) {
            if (existingNormalizedNames.includes(columnName)) {
                continue
            }

            console.info(`Adding "${info.name}" column`)
            ++lastColumn
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

            if (info.defaultFormat != null) {
                sheet.getRange(GSheetProjectSettings.firstDataRow, lastColumn, maxRows, 1)
                    .setNumberFormat(info.defaultFormat)
            }

            if (info.defaultHorizontalAlignment?.length) {
                sheet.getRange(GSheetProjectSettings.firstDataRow, lastColumn, maxRows, 1)
                    .setHorizontalAlignment(info.defaultHorizontalAlignment)
            }

            if (info.hiddenByDefault) {
                sheet.hideColumns(lastColumn)
            }

            existingNormalizedNames.push(columnName)
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
                const formulaToExpect = `
                    ={
                        "${Utils.escapeFormulaString(info.name)}";
                        ${Utils.processFormula(info.arrayFormula)}
                    }
                `
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
                    const formula = Utils.processFormula(dataValidation.getCriteriaValues()[0].toString())
                    dataValidation = dataValidation.copy()
                        .requireFormulaSatisfied(formula)
                        .build()
                }
            }
            range.setDataValidation(dataValidation)
        }

        sheet.getRange(1, 1, lastColumn, 1)
            .setHorizontalAlignment('center')
            .setFontWeight('bold')
            .setNumberFormat('')

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
    defaultFormat?: string
    defaultHorizontalAlignment?: HorizontalAlignment
    hiddenByDefault?: boolean
}

type WidthString = '#height' | '#default-height'
