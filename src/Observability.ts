class Observability {

    static reportError(message: any) {
        console.error(message)
        SpreadsheetApp.getActiveSpreadsheet().toast(message?.toString() ?? '', "Automation error")
    }

    static reportWarning(message: any) {
        console.warn(message)
    }

    static timed<T>(timerLabel: string, action: () => T, enabled?: boolean): T {
        if (enabled === false) {
            return action()
        }

        console.time(timerLabel)
        try {
            return action()
        } finally {
            console.timeEnd(timerLabel)
        }
    }

}
