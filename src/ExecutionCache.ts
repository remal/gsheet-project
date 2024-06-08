class ExecutionCache {

    private static _data = new Map<string, any>()

    static getOrCompute<T>(key: any, compute: () => T, timerLabel?: string): T {
        const stringKey = this._getStringKey(key)
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

    public static put(key: any, value: any) {
        const stringKey = this._getStringKey(key)
        this._data.set(stringKey, value)
    }

    private static _getStringKey(key: any): string {
        return JSON.stringify(key, (_, value) => {
            if (Utils.isFunction(value.getUniqueId)) {
                return value.getUniqueId()
            } else if (Utils.isFunction(value.getSheetId)) {
                return value.getSheetId()
            } else if (Utils.isFunction(value.getId)) {
                return value.getId()
            }
            return value
        })
    }

    static resetCache() {
        this._data.clear()
    }

}
