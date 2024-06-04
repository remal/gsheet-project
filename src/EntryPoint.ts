class EntryPoint {

    static entryPoint<T>(action: () => T): T {
        try {
            ExecutionCache.resetCache()
            return action()

        } catch (e) {
            console.error(e)
            SpreadsheetApp.getActiveSpreadsheet().toast(e.toString(), "Automation error")
            throw e

        } finally {
            ProtectionLocks.release()
            ProtectionLocks.releaseExpiredLocks()
        }
    }

}
