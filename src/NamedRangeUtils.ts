class NamedRangeUtils {

    static findNamedRange(rangeName: string): NamedRange | undefined {
        const namedRanges = ExecutionCache.getOrCompute('named-ranges', () => {
            const result = new Map<string, NamedRange>()
            for (const namedRange of SpreadsheetApp.getActiveSpreadsheet().getNamedRanges()) {
                const name = Utils.normalizeName(namedRange.getName())
                result.set(name, namedRange)
            }
            return result
        }, `${NamedRangeUtils.name}: ${this.findNamedRange.name}`)

        rangeName = Utils.normalizeName(rangeName)
        return namedRanges.get(rangeName)
    }

    static getNamedRange(rangeName: string): NamedRange {
        return this.findNamedRange(rangeName) ?? (() => {
            throw new Error(`"${rangeName}" named range can't be found`)
        })()
    }

    static getNamedRangeColumn(rangeName: string): Column {
        return this.getNamedRange(rangeName).getRange().getColumn()
    }

}
