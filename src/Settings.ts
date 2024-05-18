class Settings {

    static getMatrix(settingsSheet: Sheet | string, settingsScope: string): Map<string, string>[] {
        if (Utils.isString(settingsSheet)) {
            settingsSheet = SheetUtils.getSheetByName(settingsSheet)
        }

        settingsScope = Utils.normalizeName(settingsScope)

        return ExecutionCache.getOrComputeCache(['settings', 'map', settingsSheet, settingsScope], () => {
            const scopeRow = this.findScopeRow(settingsSheet, settingsScope)

            const columns: string[] = []
            const columnsValues = settingsSheet.getRange(scopeRow + 1, 1, settingsSheet.getLastColumn(), 1,).getValues()[0]
            for (const column of columnsValues) {
                const name = column.toString().trim()
                if (name.length) {
                    columns.push(name)
                } else {
                    break
                }
            }

            if (!columns.length) {
                return []
            }

            const result: Map<string, string>[] = []
            for (const row of Utils.range(scopeRow + 2, settingsSheet.getLastRow())) {
                const item = new Map<string, string>()
                const values = settingsSheet.getRange(row, 1, 1, columns.length).getValues()[0]
                for (let i = 0; i < columns.length; ++i) {
                    item[columns[i]] = values[i].toString().trim()
                }
                result.push(item)
            }
            return result
        })
    }

    static getMap(settingsSheet: Sheet | string, settingsScope: string): Map<string, string> {
        if (Utils.isString(settingsSheet)) {
            settingsSheet = SheetUtils.getSheetByName(settingsSheet)
        }

        settingsScope = Utils.normalizeName(settingsScope)

        return ExecutionCache.getOrComputeCache(['settings', 'map', settingsSheet, settingsScope], () => {
            const scopeRow = this.findScopeRow(settingsSheet, settingsScope)

            const result = new Map<string, string>()
            for (const row of Utils.range(scopeRow + 1, settingsSheet.getLastRow())) {
                const values = settingsSheet.getRange(row, 1, 1, 2).getValues()[0]
                const key = values[0].toString().trim()
                const value = values[1].toString().trim()
                if (!key.length) {
                    break
                }
                result[key] = value
            }
            return result
        })
    }

    private static findScopeRow(sheet: Sheet, scope: string): number | null {
        for (const row of Utils.range(1, sheet.getLastRow())) {
            const range = sheet.getRange(row, 1)
            if (range.getFontWeight() !== 'bold'
                || !range.isPartOfMerge()
            ) {
                continue
            }

            if (Utils.normalizeName(range.getValue()) === scope) {
                return row
            }
        }

        return null
    }

}
