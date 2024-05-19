class State {

    private static _now: number = new Date().getTime()

    static isStructureChanged(): boolean {
        const timestamp = this._loadStateTimestamp('lastStructureChange')
        return timestamp != null && this._now < timestamp
    }

    static updateLastStructureChange() {
        this._now = new Date().getTime()
        this._saveStateTimestamp('lastStructureChange', this._now)
    }


    static reset() {
        this._now = new Date().getTime()
    }


    private static _loadStateTimestamp(key: string): number | null {
        const cache = CacheService.getDocumentCache()
        if (cache == null) {
            return null
        }

        const timestamp = parseInt(cache.get(`state:${key}`) ?? '')
        return isNaN(timestamp) ? null : timestamp
    }

    private static _saveStateTimestamp(key: string, timestamp: number) {
        CacheService.getDocumentCache()?.put(`state:${key}`, timestamp.toString())
    }

}
