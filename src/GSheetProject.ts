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
    }

}
