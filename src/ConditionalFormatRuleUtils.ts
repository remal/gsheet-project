class ConditionalFormatRuleUtils {

    static extractFormula(rule: ConditionalFormatRule | ConditionalFormatRuleBuilder): string | undefined {
        const condition = rule.getBooleanCondition()
        if (condition == null) {
            return undefined
        }

        if (condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
            return undefined
        }

        return condition.getCriteriaValues()[0].toString()
    }

    static extractRequiredFormula(rule: ConditionalFormatRule | ConditionalFormatRuleBuilder): string {
        return this.extractFormula(rule) ?? (() => {
            throw new Error('Not a boolean condition with formula')
        })()
    }

}
