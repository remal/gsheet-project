class NamedRangeUtils {

    static findNamedRange(rangeName: string): NamedRange | undefined {
        rangeName = Utils.normalizeName(rangeName)
        return ExecutionCache.getOrComputeCache(['findNamedRange', rangeName], () => {
            for (const namedRange of SpreadsheetApp.getActiveSpreadsheet().getNamedRanges()) {
                const name = Utils.normalizeName(namedRange.getName())
                if (name === rangeName) {
                    return namedRange
                }
            }
            return undefined
        })
    }

    static getNamedRange(rangeName: string): NamedRange {
        return this.findNamedRange(rangeName) ?? (() => {
            throw new Error(`"${rangeName}" named range can't be found`)
        })()
    }

}
