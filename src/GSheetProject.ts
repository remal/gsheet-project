class GSheetProject {

    static reloadIssues() {
        EntryPoint.entryPoint(() => {
        })
    }

    static migrateColumns() {
        EntryPoint.entryPoint(() => {
            ProjectSheetLayout.instance.migrateColumns()
        })
    }

    static onOpen(event?: SheetsOnOpen) {
        EntryPoint.entryPoint(() => {
        })
    }

    static onChange(event?: SheetsOnChange) {
        if (!['INSERT_ROW'].includes(event?.changeType?.toString() ?? '')) {
            return
        }

        EntryPoint.entryPoint(() => {
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
        })
    }

}
