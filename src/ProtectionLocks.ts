class ProtectionLocks {

    private static readonly _allColumnsProtections = new Map<number, Protection>()
    private static readonly _allRowsProtections = new Map<number, Protection>()
    private static readonly _rowsProtections = new Map<number, Map<Row, Protection>>()

    static lockAllColumns(sheet: Sheet) {
        if (!GSheetProjectSettings.lockColumns) {
            return
        }

        const sheetId = sheet.getSheetId()
        if (this._allColumnsProtections.has(sheetId)) {
            return
        }

        Utils.timed(`${ProtectionLocks.name}: ${this.lockAllColumns.name}: ${sheet.getSheetName()}`, () => {
            const range = sheet.getRange(1, 1, 1, sheet.getMaxColumns())
            const protection = range.protect()
                .setDescription(`lock|columns|all|${new Date().getTime()}`)
                .setWarningOnly(true)
            this._allColumnsProtections.set(sheetId, protection)
        })
    }

    static lockAllRows(sheet: Sheet) {
        if (!GSheetProjectSettings.lockRows) {
            return
        }

        const sheetId = sheet.getSheetId()
        if (this._allRowsProtections.has(sheetId)) {
            return
        }

        Utils.timed(`${ProtectionLocks.name}: ${this.lockAllRows.name}: ${sheet.getSheetName()}`, () => {
            const range = sheet.getRange(1, sheet.getMaxColumns(), sheet.getMaxRows(), 1)
            const protection = range.protect()
                .setDescription(`lock|rows|all|${new Date().getTime()}`)
                .setWarningOnly(true)
            this._allRowsProtections.set(sheetId, protection)
        })
    }

    static lockRows(sheet: Sheet, rowToLock: Row) {
        if (!GSheetProjectSettings.lockRows) {
            return
        }

        if (rowToLock <= 0) {
            return
        }

        const sheetId = sheet.getSheetId()
        if (this._allRowsProtections.has(sheetId)) {
            return
        }

        if (!this._rowsProtections.has(sheetId)) {
            this._rowsProtections.set(sheetId, new Map())
        }

        const rowsProtections = this._rowsProtections.get(sheetId)!
        const maxLockedRow = Array.from(rowsProtections.keys()).reduce((prev, cur) => Math.max(prev, cur), 0)
        if (maxLockedRow < rowToLock) {
            Utils.timed(
                `${ProtectionLocks.name}: ${this.lockRows.name}: ${sheet.getSheetName()}: ${rowToLock}`,
                () => {
                    const range = sheet.getRange(1, sheet.getMaxColumns(), rowToLock, 1)
                    const protection = range.protect()
                        .setDescription(`lock|rows|${rowToLock}|${new Date().getTime()}`)
                        .setWarningOnly(true)
                    rowsProtections.set(rowToLock, protection)
                },
            )
        }
    }

    static release() {
        if (!GSheetProjectSettings.lockColumns && !GSheetProjectSettings.lockRows) {
            return
        }

        Utils.timed(`${ProtectionLocks.name}: ${this.release.name}`, () => {
            this._allColumnsProtections.forEach(protection => protection.remove())
            this._allColumnsProtections.clear()

            this._allRowsProtections.forEach(protection => protection.remove())
            this._allRowsProtections.clear()

            this._rowsProtections.forEach(protections =>
                Array.from(protections.values()).forEach(protection => protection.remove()),
            )
            this._rowsProtections.clear()
        })
    }

    static releaseExpiredLocks() {
        if (!GSheetProjectSettings.lockColumns && !GSheetProjectSettings.lockRows) {
            return
        }

        Utils.timed(`${ProtectionLocks.name}: ${this.releaseExpiredLocks.name}`, () => {
            const maxLockDurationMillis = 10 * 60 * 1000
            const minTimestamp = new Date().getTime() - maxLockDurationMillis
            SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => {
                for (const protection of sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)) {
                    const description = protection.getDescription()
                    if (!description.startsWith('lock|')) {
                        continue
                    }

                    const date = Utils.parseDate(description.split('|').slice(-1)[0])
                    if (date != null && date.getTime() < minTimestamp) {
                        console.warn(`Removing expired protection lock: ${description}`)
                        protection.remove()
                    }
                }
            })
        })
    }

}
