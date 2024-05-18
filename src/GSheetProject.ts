class GSheetProject {

    private issueIdFormatter: IssueIdFormatter
    private issueInfoLoader: IssueInfoLoader

    constructor(settings: GSheetProjectSettings) {
        this.issueIdFormatter = new IssueIdFormatter(settings);
        this.issueInfoLoader = new IssueInfoLoader(settings);
    }


    onOpen(event: SheetsOnOpen) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache()
        })
    }

    onChange(event: SheetsOnChange) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache()
        })
    }

    osEdit(event: SheetsOnEdit) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache()
            this.issueIdFormatter.formatIssueId(event.range)
            this.issueInfoLoader.loadIssueInfo(event.range)
        })
    }

    refresh() {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache()
            this.issueInfoLoader.loadAllIssueInfo()
        })
    }

}
