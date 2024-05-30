class GSheetProject {

    static reloadIssues() {
        EntryPoint.entryPoint(() => {
        })
    }

    static migrateColumns() {
        EntryPoint.entryPoint(() => {
            ProjectsSheetLayout.instance.migrateColumns()
        })
    }

    static onOpen(event?: SheetsOnOpen) {
        EntryPoint.entryPoint(() => {
            ProjectsSheetLayout.instance.migrateColumns()
        })
    }

    static onChange(event?: SheetsOnChange) {
        if (!['INSERT_ROW', 'OTHER'].includes(event?.changeType?.toString() ?? '')) {
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
