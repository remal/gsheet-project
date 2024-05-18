class GSheetProjectSettings {

    settingsSheetName: string = "Settings"

    issueColumnName: string = "Issue"
    parentIssueColumnName: string = "Parent Issue"

    issueIdsExtractor: IssueIdsExtractor = (_) => {
        throw new Error('issueIdsExtractor is not set')
    }
    issueIdDecorator: IssueIdDecorator = (id) => id
    issueIdToUrl: IssueIdToUrl = (_) => {
        throw new Error('issueIdToUrl is not set')
    }

}
