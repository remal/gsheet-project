class ExecutionCache {

    private static data = new Map<string, any>()

    static getOrComputeCache<T>(key: any, compute: () => T): T {
        const stringKey = JSON.stringify(key, (_, value) => {
            if (Utils.isFunction(value.getId)) {
                return value.getId()
            } else if (Utils.isFunction(value.getSheetId)) {
                return value.getSheetId()
            }
            return value
        })

        if (this.data.has(stringKey)) {
            return this.data.get(stringKey)
        }

        const result = compute()
        this.data.set(stringKey, result)
        return result
    }

    static resetCache() {
        this.data.clear()
    }

}
