class GSheetProject {

    private issueIdFormatter: IssueIdFormatter
    private issueInfoLoader: IssueInfoLoader

    constructor(settings: Partial<GSheetProjectSettings>) {
        const allSettings: GSheetProjectSettings = Object.assign({}, DEFAULT_GSHEET_PROJECT_SETTINGS, settings) as any
        this.issueIdFormatter = new IssueIdFormatter(allSettings);
        this.issueInfoLoader = new IssueInfoLoader(allSettings);
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
