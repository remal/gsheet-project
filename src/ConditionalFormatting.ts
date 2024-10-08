class ConditionalFormatting {

    static addConditionalFormatRule(
        range: Range,
        orderedRule: OrderedConditionalFormatRule,
        addIsFormulaRule: boolean = true,
    ) {
        if (!GSheetProjectSettings.updateConditionalFormatRules) {
            return
        }

        if ((orderedRule.order | 0) !== orderedRule.order) {
            throw new Error(`Order is not integer: ${orderedRule.order}`)
        }
        if (orderedRule.order <= 0) {
            throw new Error(`Order is <= 0: ${orderedRule.order}`)
        }

        const builder = SpreadsheetApp.newConditionalFormatRule().setRanges([range])
        orderedRule.configurer(builder)

        let formula = ConditionalFormatRuleUtils.extractRequiredFormula(builder)
        formula = Formulas.processFormula(formula).replace(/^\s*=+\s*/, '')

        const newRuleFormula = Formulas.processFormula(`=
            AND(
                ${formula},
                "GSPs"<>"${orderedRule.scope}",
                "GSPo"<>"${orderedRule.order + 0.2}"
            )
        `)
        builder.whenFormulaSatisfied(Formulas.deduplicateRowCells(newRuleFormula))
        const newRule = builder.build()
        const newRules = [newRule]

        if (addIsFormulaRule) {
            const newIsFormula = Formulas.processFormula(`=
                AND(
                    ISFORMULA(#SELF),
                    ${formula},
                    "GSPs"<>"${orderedRule.scope}",
                    "GSPo"<>"${orderedRule.order + 0.1}"
                )
            `)
            const newIsFormulaRule = newRule.copy()
                .whenFormulaSatisfied(Formulas.deduplicateRowCells(newIsFormula))
                .setItalic(true)
                .build()
            newRules.push(newIsFormulaRule)
        }

        const sheet = range.getSheet()
        let rules = sheet.getConditionalFormatRules() ?? []
        rules = rules.filter(rule =>
            !(this._extractScope(rule) === orderedRule.scope && this._extractIntOrder(rule) === orderedRule.order),
        )
        rules.push(...newRules)
        rules = rules.toSorted((r1, r2) => {
            const o1 = this._extractFloatOrder(r1) ?? 0
            const o2 = this._extractFloatOrder(r2) ?? 0
            return o1 - o2
        })
        sheet.setConditionalFormatRules(rules)
    }

    static removeConditionalFormatRulesByScope(sheet: Sheet | SheetName, scopeToRemove: string) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet)
        }

        const rules = sheet.getConditionalFormatRules() ?? []
        const filteredRules = rules.filter(rule =>
            this._extractScope(rule) !== scopeToRemove,
        )
        if (filteredRules.length !== rules.length) {
            sheet.setConditionalFormatRules(filteredRules)
        }
    }

    static removeDuplicateConditionalFormatRules(sheet?: Sheet | SheetName) {
        if (sheet == null) {
            SpreadsheetApp.getActiveSpreadsheet().getSheets()
                .filter(sheet => SheetUtils.isGridSheet(sheet))
                .forEach(sheet => this.removeDuplicateConditionalFormatRules(sheet))
            return
        }

        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet)
        }

        const rules = sheet.getConditionalFormatRules() ?? []
        const filteredRules = rules.filter(Utils.distinctBy(rule =>
            JSON.stringify(Utils.toJsonObject(rule)),
        ))
        if (filteredRules.length !== rules.length) {
            sheet.setConditionalFormatRules(filteredRules)
        }
    }

    static combineConditionalFormatRules(sheet?: Sheet | SheetName) {
        if (sheet == null) {
            SpreadsheetApp.getActiveSpreadsheet().getSheets()
                .filter(sheet => SheetUtils.isGridSheet(sheet))
                .forEach(sheet => this.combineConditionalFormatRules(sheet))
            return
        }

        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet)
        }

        const rules = sheet.getConditionalFormatRules() ?? []
        if (rules.length <= 1) {
            return
        }
        const originalRules = [...rules]

        const isMergeableRule = (rule: ConditionalFormatRule): boolean => {
            const ranges = rule.getRanges()
            const firstRange = ranges.shift()
            if (firstRange == null) {
                return false
            }

            for (const range of ranges) {
                if (range.getColumn() !== firstRange.getColumn()
                    || range.getNumColumns() !== firstRange.getNumColumns()
                ) {
                    return false
                }
            }

            return true
        }
        const getRuleKey = (rule: ConditionalFormatRule): string => {
            const jsonObject = Utils.toJsonObject(rule)
            delete jsonObject['ranges']

            const ranges = rule.getRanges()
            const firstRange = ranges.shift()!
            jsonObject['columns'] = Array.from(Utils.range(
                firstRange.getColumn(),
                firstRange.getColumn() + firstRange.getNumColumns() - 1,
            ))
            return JSON.stringify(jsonObject)
        }

        for (let index = 0; index < rules.length - 1; ++index) {
            const rule = rules[index]
            if (!isMergeableRule(rule)) {
                continue
            }

            const ruleKey = getRuleKey(rule)

            let similarRule: ConditionalFormatRule | null = null
            for (let otherIndex = index + 1; otherIndex < rules.length; ++otherIndex) {
                const otherRule = rules[otherIndex]
                if (!isMergeableRule(otherRule)) {
                    continue
                }

                const otherRuleKey = getRuleKey(otherRule)
                if (otherRuleKey === ruleKey) {
                    similarRule = otherRule
                    rules.splice(otherIndex, 1)
                    break
                }
            }
            if (similarRule == null) {
                continue
            }

            const ranges = [...rule.getRanges(), ...similarRule.getRanges()]
            let newRanges = [...ranges]
            newRanges = newRanges.toSorted((r1, r2) => {
                const row1 = r1.getRow()
                const row2 = r2.getRow()
                if (row1 === row2) {
                    return r2.getNumRows() - r1.getNumRows()
                }
                return row1 - row2
            })
            for (let rangeIndex = 0; rangeIndex < newRanges.length - 1; ++rangeIndex) {
                let range = newRanges[rangeIndex]
                let firstRow = range.getRow()
                let lastRow = firstRow + range.getNumRows() - 1

                for (let nextRangeIndex = rangeIndex + 1; nextRangeIndex < newRanges.length; ++nextRangeIndex) {
                    const nextRange = newRanges[nextRangeIndex]
                    const nextFirstRow = nextRange.getRow()
                    if (nextFirstRow <= lastRow) {
                        const nextLastRow = nextFirstRow + nextRange.getNumRows() - 1
                        firstRow = Math.min(firstRow, nextFirstRow)
                        lastRow = Math.max(lastRow, nextLastRow)
                        lastRow = Math.min(lastRow, SheetUtils.getMaxRows(sheet))
                        newRanges[rangeIndex] = range = range.getSheet().getRange(
                            firstRow,
                            range.getColumn(),
                            lastRow - firstRow + 1,
                            range.getNumColumns(),
                        )
                        newRanges.splice(nextRangeIndex, 1)
                        --nextRangeIndex
                    }
                }
            }

            console.warn([
                ConditionalFormatting.name,
                `Combining ${ranges.map(it => it.getA1Notation())} into ${newRanges.map(it => it.getA1Notation())}`,
                ruleKey,
            ].join(': '))
            rules[index] = rule.copy().setRanges(newRanges).build()
        }

        if (rules.length !== originalRules.length) {
            sheet.setConditionalFormatRules(rules)
        }
    }

    private static _extractScope(rule: ConditionalFormatRule | string): string | undefined {
        if (!Utils.isString(rule)) {
            const formula = ConditionalFormatRuleUtils.extractFormula(rule)
            if (formula == null) {
                return undefined
            }

            rule = formula
        }

        const match = rule.match(/"GSPs"\s*<>\s*"([^"]*)"/)
        if (match) {
            return match[1]
        }

        return undefined
    }

    private static _extractIntOrder(rule: ConditionalFormatRule | string): number | undefined {
        if (!Utils.isString(rule)) {
            const formula = ConditionalFormatRuleUtils.extractFormula(rule)
            if (formula == null) {
                return undefined
            }

            rule = formula
        }

        const match = rule.match(/"GSPo"\s*<>\s*"(\d+)(\.\d*)?"/)
        if (match) {
            return parseInt(match[1])
        }

        return undefined
    }

    private static _extractFloatOrder(rule: ConditionalFormatRule | string): number | undefined {
        if (!Utils.isString(rule)) {
            const formula = ConditionalFormatRuleUtils.extractFormula(rule)
            if (formula == null) {
                return undefined
            }

            rule = formula
        }

        const match = rule.match(/"GSPo"\s*<>\s*"(\d+(\.\d*)?)"/)
        if (match) {
            return parseFloat(match[1])
        }

        return undefined
    }

    private static _ruleKey(rule: ConditionalFormatRule): string {
        const jsonObject = Utils.toJsonObject(rule)
        delete jsonObject['ranges']
        return JSON.stringify(jsonObject)
    }

}

type ConditionalFormatRuleConfigurer = (builder: ConditionalFormatRuleBuilder) => void

interface OrderedConditionalFormatRule {
    scope: string
    order: number
    configurer: ConditionalFormatRuleConfigurer
}
