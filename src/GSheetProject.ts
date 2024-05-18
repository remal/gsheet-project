class GSheetProject {

    static reloadIssues() {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache()
            IssueInfoLoader.loadAllIssueInfo()
        })
    }

    static onOpen(event?: SheetsOnOpen) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache()
        })
    }

    static onChange(event?: SheetsOnChange) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache()
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
            ExecutionCache.resetCache()
            if (range != null) {
                IssueIdFormatter.formatIssueId(range)
                HierarchyFormatter.formatHierarchy(range)
                IssueInfoLoader.loadIssueInfo(range)
            }
        })
    }

}
