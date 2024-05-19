class GSheetProject {

    static reloadIssues() {
        Utils.entryPoint(() => {
            IssueLoader.loadAllIssues()
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
        })
    }

    static onEdit(event?: SheetsOnEdit) {
        this.onEditRange(event?.range)
    }

    static onFormSubmit(event: SheetsOnFormSubmit) {
        this.onEditRange(event?.range)
    }

    private static onEditRange(range?: Range) {
        Utils.entryPoint(() => {
            if (range != null) {
                IssueIdFormatter.formatIssueId(range)
                HierarchyFormatter.formatHierarchy(range)
                IssueLoader.loadIssues(range)
            }
        })
    }

}
