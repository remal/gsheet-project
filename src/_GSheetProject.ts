class GSheetProject {

    static reloadIssues() {
        EntryPoint.entryPoint(() => {
        })
    }

    static migrateColumns() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateColumns()
        })
    }

    static refreshEverything() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateColumnsIfNeeded()

            const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName)
            const range = sheet.getRange(
                GSheetProjectSettings.firstDataRow,
                1,
                Math.max(sheet.getLastRow() - GSheetProjectSettings.firstDataRow + 1, 1),
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
            SheetLayouts.migrateColumnsIfNeeded()
        })
    }


    static onChange(event?: SheetsOnChange) {
        const changeType = event?.changeType?.toString()
        if (changeType === 'INSERT_ROW') {
            this._onInsertRow()
        } else if (changeType === 'REMOVE_COLUMN') {
            this._onRemoveColumn()
        }
    }

    private static _onInsertRow() {
        EntryPoint.entryPoint(() => {
        })
    }

    private static _onRemoveColumn() {
        this.migrateColumns()
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
            DoneLogic.executeDoneLogic(range)
            DefaultFormulas.insertDefaultFormulas(range)
            IssueHierarchyFormatter.formatHierarchy(range)
        })
    }

}
