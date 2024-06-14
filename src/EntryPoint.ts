class EntryPoint {

    private static _isInEntryPoint: boolean = false

    static entryPoint<T>(action: () => T): T {
        if (this._isInEntryPoint) {
            return action()
        }

        try {
            this._isInEntryPoint = true
            ExecutionCache.resetCache()
            return action()

        } catch (e) {
            Observability.reportError(e)
            throw e

        } finally {
            ProtectionLocks.release()
            ProtectionLocks.releaseExpiredLocks()
            this._isInEntryPoint = false
        }
    }

}
