class State {

    private static now: number = new Date().getTime()

    static isStructureChanged(): boolean {
        const timestamp = this.loadStateTimestamp('lastStructureChange')
        return !isNaN(timestamp) && this.now < timestamp
    }

    static updateLastStructureChange() {
        this.now = new Date().getTime()
        this.saveStateTimestamp('lastStructureChange', this.now)
    }


    static reset() {
        this.now = new Date().getTime()
    }


    private static loadStateTimestamp(key: string): number {
        return parseInt(CacheService.getDocumentCache().get(`state:${key}`))
    }

    private static saveStateTimestamp(key: string, timestamp: number) {
        CacheService.getDocumentCache().put(`state:${key}`, timestamp.toString())
    }

}
