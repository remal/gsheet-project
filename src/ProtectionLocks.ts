class ProtectionLocks {

    private static _columnsProtections = new Map<number, Protection>()
    private static _rowsProtections = new Map<number, Protection>()

    static lockColumnsWithProtection(sheet: Sheet) {
        const sheetId = sheet.getSheetId()
        if (this._columnsProtections.has(sheetId)) {
            return
        }

        const range = sheet.getRange(1, 1, 1, sheet.getMaxColumns())
        const protection = range.protect()
            .setDescription(`lock|columns|${new Date().getTime()}`)
            .setWarningOnly(true)
        this._columnsProtections.set(sheetId, protection)
    }

    static lockRowsWithProtection(sheet: Sheet) {
        const sheetId = sheet.getSheetId()
        if (this._rowsProtections.has(sheetId)) {
            return
        }

        const range = sheet.getRange(1, sheet.getMaxColumns(), sheet.getMaxRows(), 1)
        const protection = range.protect()
            .setDescription(`lock|rows|${new Date().getTime()}`)
            .setWarningOnly(true)
        this._rowsProtections.set(sheetId, protection)
    }

    static release() {
        this._columnsProtections.forEach(protection => protection.remove())
        this._columnsProtections.clear()

        this._rowsProtections.forEach(protection => protection.remove())
        this._rowsProtections.clear()
    }

    static releaseExpiredLocks() {
        const maxLockDurationMillis = 10 * 60 * 1000
        const minTimestamp = new Date().getTime() - maxLockDurationMillis
        SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => {
            for (const protection of sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)) {
                const description = protection.getDescription()
                if (!description.startsWith('lock|')) {
                    continue
                }

                const dateString = description.split('|').slice(-1)[0]
                try {
                    const date = Number.isNaN(dateString)
                        ? new Date(dateString)
                        : new Date(parseFloat(dateString))
                    if (date.getTime() < minTimestamp) {
                        console.log(`Removing expired protection lock: ${description}`)
                        protection.remove()
                    }
                } catch (_) {
                    // do nothing
                }
            }
        })
    }

}
