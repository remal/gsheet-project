class Settings {

    static getMatrix(settingsScope: string): Map<string, string>[] {
        const settingsSheet = SheetUtils.getSheetByName(GSheetProjectSettings.settingsSheetName)
        settingsScope = Utils.normalizeName(settingsScope)
        return ExecutionCache.getOrComputeCache(['settings', 'matrix', settingsScope], () => {
            const scopeRow = this._findScopeRow(settingsSheet, settingsScope)
            if (scopeRow == null) {
                throw new Error(`Settings with "${settingsScope}" can't be found`)
            }

            const columns: string[] = []
            const columnsValues = settingsSheet
                .getRange(scopeRow + 1, 1, 1, settingsSheet.getLastColumn())
                .getValues()[0]
            for (const column of columnsValues) {
                const name = Utils.toLowerCamelCase(column.toString().trim())
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
                    let value = values[i].toString().trim()
                    item.set(columns[i], value)
                }
                const areAllValuesEmpty = Array.from(item.values()).every(value => !value.length)
                if (areAllValuesEmpty) {
                    break
                }
                result.push(item)
            }

            return result
        })
    }

    static getMap(settingsScope: string): Map<string, string> {
        const settingsSheet = SheetUtils.getSheetByName(GSheetProjectSettings.settingsSheetName)
        settingsScope = Utils.normalizeName(settingsScope)
        return ExecutionCache.getOrComputeCache(['settings', 'map', settingsScope], () => {
            const scopeRow = this._findScopeRow(settingsSheet, settingsScope)
            if (scopeRow == null) {
                throw new Error(`Settings with "${settingsScope}" can't be found`)
            }

            const result = new Map<string, string>()
            for (const row of Utils.range(scopeRow + 1, settingsSheet.getLastRow())) {
                const values = settingsSheet.getRange(row, 1, 1, 2).getValues()[0]
                const key = Utils.toLowerCamelCase(values[0].toString().trim())
                if (!key.length) {
                    break
                }
                const value = values[1].toString().trim()
                result.set(key, value)
            }
            return result
        })
    }

    private static _findScopeRow(sheet: Sheet, scope: string): number | null {
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
