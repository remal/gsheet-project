class GSheetProject {

    static reloadIssues() {
        EntryPoint.entryPoint(() => {
        })
    }

    static migrate() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrate()
        })
    }

    static refreshEverything() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateIfNeeded()

            const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName)
            const range = sheet.getRange(
                GSheetProjectSettings.firstDataRow,
                1,
                Math.max(SheetUtils.getLastRow(sheet) - GSheetProjectSettings.firstDataRow + 1, 1),
                sheet.getLastColumn(),
            )
            this._onEditRange(range)
        })
    }

    static cleanup() {
        EntryPoint.entryPoint(() => {
            ProtectionLocks.releaseExpiredLocks()
        })
    }


    static onOpen(event?: SheetsOnOpen) {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateIfNeeded()
        })
    }


    static onChange(event?: SheetsOnChange) {
        const changeType = event?.changeType?.toString()
        if (changeType === 'INSERT_ROW') {
            this._onInsertRow()
        } else if (changeType === 'INSERT_COLUMN') {
            this._onInsertColumn()
        } else if (changeType === 'REMOVE_COLUMN') {
            this._onRemoveColumn()
        }
    }

    private static _onInsertRow() {
        EntryPoint.entryPoint(() => {
            CommonFormatter.applyCommonFormatsToAllSheets()
        })
    }

    private static _onInsertColumn() {
        EntryPoint.entryPoint(() => {
            CommonFormatter.applyCommonFormatsToAllSheets()
        })
    }

    private static _onRemoveColumn() {
        this.migrate()
    }


    static onEdit(event?: SheetsOnEdit) {
        this._onEditRange(event?.range)
    }

    static onFormSubmit(event: SheetsOnFormSubmit) {
        this._onEditRange(event?.range)
    }

    private static _onEditRange(range?: Range) {
        if (range == null) {
            return
        }

        EntryPoint.entryPoint(() => {
            //Utils.timed(`Done logic`, () => DoneLogic.executeDoneLogic(range))
            Utils.timed(`Default formulas`, () => DefaultFormulas.insertDefaultFormulas(range))
            Utils.timed(`Issue hierarchy`, () => IssueHierarchyFormatter.formatHierarchy(range))
        })
    }

}
