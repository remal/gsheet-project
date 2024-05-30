class EntryPoint {

    static entryPoint<T>(action: () => T): T {
        try {
            ExecutionCache.resetCache()
            return action()

        } catch (e) {
            console.error(e)
            throw e

        } finally {
            ProtectionLocks.release()
            ProtectionLocks.releaseExpiredLocks()
        }
    }

}
