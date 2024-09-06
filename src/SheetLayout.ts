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
        return `${this._documentFlagPrefix}$$$HASH$$$:${GSheetProjectSettings.computeStringSettingsHash()}:${this.sheet.getMaxRows()}`
    }

    migrateIfNeeded(): boolean {
        if (DocumentFlags.isSet(this._documentFlag)) {
            console.log(`Skipping '${this.sheet.getSheetName()}' sheet migration`)
            return false
        }

        this.migrate()
        return true
    }

    migrate() {
        const sheet = this.sheet
        console.warn(`Migrating '${sheet.getSheetName()}' sheet`)


        const conditionalFormattingScope = `layout:${this.constructor?.name || Utils.normalizeName(this.sheetName)}`
        let conditionalFormattingOrder = 0
        ConditionalFormatting.removeConditionalFormatRulesByScope(sheet, 'layout')
        ConditionalFormatting.removeConditionalFormatRulesByScope(sheet, conditionalFormattingScope)


        const columns = this.columns.reduce(
            (map, info) => {
                map.set(Utils.normalizeName(info.name), info)
                return map
            },
            new Map<string, ColumnInfo>(),
        )
        if (!columns.size) {
            DocumentFlags.set(this._documentFlag)
            DocumentFlags.cleanupByPrefix(this._documentFlagPrefix)
            return
        }


        ProtectionLocks.lockAllColumns(sheet)

        let lastColumn = SheetUtils.getLastColumn(sheet)
        const maxRows = SheetUtils.getMaxRows(sheet)
        const existingNormalizedNames = sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn)
            .getValues()[0]
            .map(it => it?.toString())
            .map(it => it?.length ? Utils.normalizeName(it) : '')
        for (const [columnName, info] of columns.entries()) {
            const existingIndex = existingNormalizedNames.indexOf(columnName)
            if (existingIndex >= 0) {
                continue
            }

            console.info(`Adding "${info.name}" column`)
            ++lastColumn
            const titleRange = sheet.getRange(GSheetProjectSettings.titleRow, lastColumn)
                .setValue(info.name)

            ExecutionCache.resetCache()

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
                const formulaToExpect = Formulas.processFormula(`=
                    {
                        "${Formulas.escapeFormulaString(info.name).replaceAll(/[\r\n]+/g, '"&CHAR(10)&"')}";
                        ${Formulas.processFormula(info.arrayFormula)}
                    }
                `)
                const formula = existingFormulas.get()[index]
                if (formula !== formulaToExpect) {
                    sheet.getRange(GSheetProjectSettings.titleRow, column)
                        .setFormula(formulaToExpect)
                }
            }

            const range = sheet.getRange(
                GSheetProjectSettings.firstDataRow,
                column,
                maxRows - GSheetProjectSettings.firstDataRow + 1,
                1,
            )
            if (info.rangeName?.length) {
                SpreadsheetApp.getActiveSpreadsheet().setNamedRange(info.rangeName, range)
            }


            let dataValidation: (DataValidation | null) = info.dataValidation != null
                ? info.dataValidation()
                : null
            if (dataValidation != null) {
                if (dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CUSTOM_FORMULA) {
                    const formula = Formulas.processFormula(dataValidation.getCriteriaValues()[0].toString())
                    dataValidation = dataValidation.copy()
                        .requireFormulaSatisfied(formula)
                        .build()
                }
                range.setDataValidation(dataValidation)
            }


            function getConfigurerOf(
                configurer: ConditionalFormatRuleConfigurer | ParametrizedConditionalFormatRuleConfigurer,
            ): ConditionalFormatRuleConfigurer {
                const untyped = configurer as any
                return untyped.configurer != null ? untyped.configurer : untyped
            }

            function isMergeWithPrevious(
                configurer: ConditionalFormatRuleConfigurer | ParametrizedConditionalFormatRuleConfigurer,
            ) {
                const untyped = configurer as any
                return !!untyped.mergeWithPrevious
            }

            const allConditionalFormats = info.conditionalFormats?.filter(it => it != null) ?? []
            const mergedConditionalFormats: ConditionalFormatRuleConfigurer[] = []
            const conditionalFormats: ConditionalFormatRuleConfigurer[] = []
            for (let index = allConditionalFormats.length - 1; 0 <= index; --index) {
                const conditionalFormat = allConditionalFormats[index]
                const configurer = getConfigurerOf(conditionalFormat)
                conditionalFormats.unshift(configurer)

                if (isMergeWithPrevious(conditionalFormat)) {
                    for (let prevIndex = index - 1; 0 <= prevIndex; --prevIndex) {
                        const prevConditionalFormat = allConditionalFormats[prevIndex]
                        const prevConfigurer = getConfigurerOf(prevConditionalFormat)
                        const mergedConfigurer: ConditionalFormatRuleConfigurer = builder => {
                            configurer(builder)
                            const formula = ConditionalFormatRuleUtils.extractRequiredFormula(builder)

                            prevConfigurer(builder)
                            const prevFormula = ConditionalFormatRuleUtils.extractRequiredFormula(builder)

                            let combinedFormula = '=AND(' + [
                                formula,
                                prevFormula,
                            ].map(it => it.replace(/^\s*=+\s*/, '')) + ')'
                            builder.whenFormulaSatisfied(combinedFormula)
                        }
                        mergedConditionalFormats.unshift(mergedConfigurer)
                    }
                }
            }
            conditionalFormats.unshift(...mergedConditionalFormats)

            conditionalFormats.forEach(configurer => {
                const fullRule = {
                    scope: conditionalFormattingScope,
                    order: ++conditionalFormattingOrder,
                    configurer,
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

interface ColumnInfo {
    name: ColumnName
    arrayFormula?: string
    rangeName?: RangeName
    dataValidation?: () => (DataValidation | null)
    conditionalFormats?: (ConditionalFormatRuleConfigurer | ParametrizedConditionalFormatRuleConfigurer | null | undefined)[]
    defaultTitleFontSize?: number
    defaultWidth?: number | WidthString
    defaultFormat?: string
    defaultHorizontalAlignment?: HorizontalAlignment
    hiddenByDefault?: boolean
}

interface ParametrizedConditionalFormatRuleConfigurer {
    mergeWithPrevious?: boolean
    configurer: ConditionalFormatRuleConfigurer
}

type WidthString = '#height' | '#default-height'
