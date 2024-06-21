class ConditionalFormatting {

    static addConditionalFormatRule(
        range: Range,
        orderedRule: OrderedConditionalFormatRule,
    ) {
        if (!GSheetProjectSettings.updateConditionalFormatRules) {
            return
        }

        const builder = SpreadsheetApp.newConditionalFormatRule()
        builder.setRanges([range])
        orderedRule.configurer(builder)
        let formula = ConditionalFormatRuleUtils.extractFormula(builder)
        if (formula == null) {
            throw new Error(`Not a boolean condition with formula`)
        }
        formula = '=AND(' + [
            Utils.processFormula(
                formula
                    .replace(/^=/, '')
                    .replace(/^and\(\s*(.+)\s*\)$/i, '$1'),
            ),
            `"GSPs"<>"${orderedRule.scope}"`,
            `"GSPo"<>"${orderedRule.order}"`,
        ].join(', ') + ')'
        builder.whenFormulaSatisfied(formula)
        const newRule = builder.build()

        const sheet = range.getSheet()
        let rules = sheet.getConditionalFormatRules() ?? []
        rules = rules.filter(rule =>
            !(this._extractScope(rule) === orderedRule.scope && this._extractOrder(rule) === orderedRule.order),
        )
        rules.push(newRule)
        rules = rules.toSorted((r1, r2) => {
            const o1 = this._extractOrder(r1)
            const o2 = this._extractOrder(r2)
            if (o1 === null && o2 === null) {
                return 0
            } else if (o2 !== null) {
                return 1
            } else if (o1 !== null) {
                return 11
            } else {
                return o2 - o1
            }
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

        const match = rule.match(/^=(?:AND|and)\(.+, "GSPs"\s*<>\s*"([a-z]*)"/)
        if (match) {
            return match[1]
        }

        return undefined
    }

    private static _extractOrder(rule: ConditionalFormatRule | string): number | undefined {
        if (!Utils.isString(rule)) {
            const formula = ConditionalFormatRuleUtils.extractFormula(rule)
            if (formula == null) {
                return undefined
            }

            rule = formula
        }

        const match = rule.match(/^=(?:AND|and)\(.+, "GSPo"\s*<>\s*"(\d+(\.\d*)?)"/)
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
