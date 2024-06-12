abstract class SheetLayout {

    protected abstract get sheetName(): SheetName

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


        ConditionalFormatting.removeConditionalFormatRulesByScope(sheet, 'layout')


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

        const columnByKey = new Map<string, { columnNumber: Column, info: ColumnInfo }>()

        let lastColumn = SheetUtils.getLastColumn(sheet)
        const maxRows = sheet.getMaxRows()
        const existingNormalizedNames = sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn)
            .getValues()[0]
            .map(it => it?.toString())
            .map(it => it?.length ? Utils.normalizeName(it) : '')
        for (const [columnName, info] of columns.entries()) {
            const existingIndex = existingNormalizedNames.indexOf(columnName)
            if (existingIndex >= 0) {
                if (info.key?.length) {
                    const columnNumber = existingIndex + 1
                    columnByKey.set(info.key, {columnNumber, info})
                }
                continue
            }

            console.info(`Adding "${info.name}" column`)
            ++lastColumn
            const titleRange = sheet.getRange(GSheetProjectSettings.titleRow, lastColumn)
                .setValue(info.name)

            ExecutionCache.resetCache()

            if (info.key?.length) {
                const columnNumber = lastColumn
                columnByKey.set(info.key, {columnNumber, info})
            }

            if (info.defaultTitleFontSize != null && info.defaultTitleFontSize > 0) {
                titleRange.setFontSize(info.defaultTitleFontSize)
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
        SheetUtils.setLastColumn(sheet, lastColumn)

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


            const processFormula = (formula: string): string => {
                formula = Utils.processFormula(formula)
                formula = formula.replaceAll(/#COLUMN_CELL\(([^)]+)\)/g, (_, key) => {
                    const columnNumber = columnByKey.get(key)?.columnNumber
                    if (columnNumber == null) {
                        throw new Error(`Column with key '${key}' can't be found`)
                    }
                    return sheet.getRange(GSheetProjectSettings.firstDataRow, columnNumber).getA1Notation()
                })
                formula = formula.replaceAll(/#COLUMN_CELL\b/g, () => {
                    return range.getCell(1, 1).getA1Notation()
                })
                return formula
            }


            let dataValidation: (DataValidation | null) = info.dataValidation != null
                ? info.dataValidation()
                : null
            if (dataValidation != null) {
                if (dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CUSTOM_FORMULA) {
                    const formula = processFormula(dataValidation.getCriteriaValues()[0].toString())
                    dataValidation = dataValidation.copy()
                        .requireFormulaSatisfied(formula)
                        .build()
                }
            }
            range.setDataValidation(dataValidation)

            info.conditionalFormats?.forEach(rule => {
                const originalConfigurer = rule.configurer
                rule.configurer = builder => {
                    originalConfigurer(builder)
                    const formula = ConditionalFormatRuleUtils.extractFormula(builder)
                    if (formula != null) {
                        builder.whenFormulaSatisfied(processFormula(formula))
                    }
                    return builder
                }
                const fullRule = {
                    scope: 'layout',
                    ...rule,
                }
                ConditionalFormatting.addConditionalFormatRule(range, fullRule)
            })
        }

        sheet.getRange('1:1')
            .setHorizontalAlignment('center')
            .setFontWeight('bold')
            .setFontLine('none')
            .setNumberFormat('')

        DocumentFlags.set(this._documentFlag)
        DocumentFlags.cleanupByPrefix(this._documentFlagPrefix)

        const waitForAllDataExecutionsCompletion = SpreadsheetApp.getActiveSpreadsheet()['waitForAllDataExecutionsCompletion']
        if (Utils.isFunction(waitForAllDataExecutionsCompletion)) {
            try {
                waitForAllDataExecutionsCompletion(5)
            } catch (e) {
                console.warn(e)
            }
        }
    }

}

type LayoutOrderedConditionalFormatRule = Omit<OrderedConditionalFormatRule, 'scope'>

interface ColumnInfo {
    key?: string
    name: ColumnName
    arrayFormula?: string
    rangeName?: RangeName
    dataValidation?: () => (DataValidation | null)
    conditionalFormats?: LayoutOrderedConditionalFormatRule[]
    defaultTitleFontSize?: number
    defaultWidth?: number | WidthString
    defaultFormat?: string
    defaultHorizontalAlignment?: HorizontalAlignment
    hiddenByDefault?: boolean
}

type WidthString = '#height' | '#default-height'
