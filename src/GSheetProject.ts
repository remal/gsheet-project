class GSheetProject {

    static reloadIssues() {
        EntryPoint.entryPoint(() => {
        })
    }

    static migrateColumns() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateColumnsIfNeeded()
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
            IssueHierarchyFormatter.formatHierarchyForAllIssues()
        })
    }

    private static _onRemoveColumn() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateColumns()
        })
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
            IssueHierarchyFormatter.formatHierarchy(range)
        })
    }

}
