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

}

interface OrderedConditionalFormatRule {
    scope: string
    order: number
    configurer: (builder: ConditionalFormatRuleBuilder) => void
}
