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

}
