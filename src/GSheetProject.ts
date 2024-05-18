class GSheetProject {

    private settings: GSheetProjectSettings

    constructor(settings: GSheetProjectSettings) {
        this.settings = settings;
    }


    onOpen(event: SheetsOnOpen) {
        ExecutionCache.resetCache()
    }

    onChange(event: SheetsOnChange) {
        ExecutionCache.resetCache()
    }

    osEdit(event: SheetsOnEdit) {
        ExecutionCache.resetCache()
        StyleIssueId.formatIssueId(
            event.range,
            this.settings.issueIdsExtractor,
            this.settings.issueIdDecorator,
            this.settings.issueIdToUrl,
            this.settings.issueColumnName,
            this.settings.parentIssueColumnName,
        )
    }

}
