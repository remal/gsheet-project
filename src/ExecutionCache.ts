class ExecutionCache {

    private static _data = new Map<string, any>()

    static getOrComputeCache<T>(key: any, compute: () => T, timerLabel?: string): T {
        const stringKey = JSON.stringify(key, (_, value) => {
            if (Utils.isFunction(value.getUniqueId)) {
                return value.getUniqueId()
            } else if (Utils.isFunction(value.getSheetId)) {
                return value.getSheetId()
            } else if (Utils.isFunction(value.getId)) {
                return value.getId()
            }
            return value
        })

        if (this._data.has(stringKey)) {
            return this._data.get(stringKey)
        }


        if (timerLabel?.length) {
            console.time(timerLabel)
        }

        let result: T
        try {
            result = compute()

        } finally {
            if (timerLabel?.length) {
                console.timeEnd(timerLabel)
            }
        }


        this._data.set(stringKey, result)
        return result
    }

    static resetCache() {
        this._data.clear()
    }

}
