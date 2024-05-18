class GSheetProjectSettings {

    firstDataRow: number = 2

    settingsSheetName: string = "Settings"

    issueIdColumnName: string = "Issue"
    parentIssueIdColumnName: string = "Parent Issue"

    isDoneColumnName?: string
    idDoneCalculator: IssueIsDoneCalculator = () => {
        throw new Error('idDoneCalculator is not set')
    }

    stringFields: Record<string, IssueStringFieldGetter> = {}
    booleanFields: Record<string, IssueBooleanFieldGetter> = {}

    childIssueMetrics: IssueMetric[] = []
    blockerIssueMetrics: IssueMetric[] = []

    issueIdsExtractor: IssueIdsExtractor = () => {
        throw new Error('issueIdsExtractor is not set')
    }
    issueIdDecorator: IssueIdDecorator = (id) => id
    issueIdToUrl: IssueIdToUrl = () => {
        throw new Error('issueIdToUrl is not set')
    }
    issueIdsToUrl?: IssueIdsToUrl = null

    issuesLoader: IssuesLoader = () => {
        throw new Error('issuesLoader is not set')
    }
    childIssuesLoader: IssuesLoader = () => {
        throw new Error('childIssuesLoader is not set')
    }
    blockerIssuesLoader: IssuesLoader = () => {
        throw new Error('blockerIssuesLoader is not set')
    }

    issueIdGetter: IssueStringFieldGetter = () => {
        throw new Error('issueIdGetter is not set')
    }

}
