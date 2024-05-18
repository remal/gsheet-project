class GSheetProjectSettings {

    settingsSheetName: string = "Settings"

    issueColumnName: string = "Issue"
    parentIssueColumnName: string = "Parent Issue"

    issueIdsExtractor: (string: string) => string[] = (_) => {
        throw new Error('issueIdsExtractor is not set')
    }
    issueIdToLink: (id: string) => string = (_) => {
        throw new Error('issueIdToLink is not set')
    }

}
