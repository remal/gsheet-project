class GSheetProject {

    static reloadIssues() {
        Utils.entryPoint(() => {
            IssueLoader.loadAllIssues()
        })
    }

    static recalculateSchedule() {
        Utils.entryPoint(() => {
            Schedule.recalculateAllSchedules()
        })
    }

    static onOpen(event?: SheetsOnOpen) {
        Utils.entryPoint(() => {
        })
    }

    static onChange(event?: SheetsOnChange) {
        Utils.entryPoint(() => {
            State.updateLastStructureChange()
            HierarchyFormatter.formatAllHierarchy()
            Schedule.recalculateAllSchedules()
        })
    }

    static onEdit(event?: SheetsOnEdit) {
        this._onEditRange(event?.range)
    }

    static onFormSubmit(event: SheetsOnFormSubmit) {
        this._onEditRange(event?.range)
    }

    private static _onEditRange(range?: Range) {
        Utils.entryPoint(() => {
            if (range != null) {
                IssueIdFormatter.formatIssueId(range)
                HierarchyFormatter.formatHierarchy(range)
                IssueLoader.loadIssues(range)
                Schedule.recalculateSchedule(range)
            }
        })
    }

}
