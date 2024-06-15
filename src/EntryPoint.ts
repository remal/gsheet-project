class EntryPoint {

    private static _isInEntryPoint: boolean = false

    static entryPoint<T>(action: () => T, useLocks?: boolean): T {
        if (this._isInEntryPoint) {
            return action()
        }

        let lock: Lock | null = null
        if (useLocks ?? GSheetProjectSettings.useLockService) {
            lock = LockService.getDocumentLock()
            lock.waitLock(GSheetProjectSettings.lockTimeout)
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
            lock?.releaseLock()
            this._isInEntryPoint = false
        }
    }

}
