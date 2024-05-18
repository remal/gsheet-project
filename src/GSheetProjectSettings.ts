class GSheetProjectSettings {

    static firstDataRow: number = 2

    static settingsSheetName: string = "Settings"

    static issueIdColumnName: string = "Issue"
    static parentIssueIdColumnName: string = "Parent Issue"

    static isDoneColumnName?: string = "Done"

    static issueIdsExtractor: IssueIdsExtractor = () => Utils.throwNotConfigured('issueIdsExtractor')
    static issueIdDecorator: IssueIdDecorator = () => Utils.throwNotConfigured('issueIdDecorator')
    static issueIdToUrl: IssueIdToUrl = () => Utils.throwNotConfigured('issueIdToUrl')
    static issueIdsToUrl?: IssueIdsToUrl = () => Utils.throwNotConfigured('issueIdsToUrl')

    static issuesLoader: IssuesLoader = () => Utils.throwNotConfigured('issuesLoader')
    static childIssuesLoader: IssuesLoader = () => Utils.throwNotConfigured('childIssuesLoader')
    static blockerIssuesLoader: IssuesLoader = () => Utils.throwNotConfigured('blockerIssuesLoader')

    static issueIdGetter: IssueStringFieldGetter = () => Utils.throwNotConfigured('issueIdGetter')

    static idDoneCalculator: IssueAggregateBooleanFieldGetter = () => Utils.throwNotConfigured('idDoneCalculator')

    static stringFields: Record<string, IssueStringFieldGetter> = {}
    static booleanFields: Record<string, IssueBooleanFieldGetter> = {}

    static childIssueMetrics: IssueMetric[] = []
    static blockerIssueMetrics: IssueMetric[] = []

}
