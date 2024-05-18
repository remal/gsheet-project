const DEFAULT_GSHEET_PROJECT_SETTINGS: Partial<GSheetProjectSettings> = {

    firstDataRow: 2,

    settingsSheetName: "Settings",

    issueIdColumnName: "Issue",
    parentIssueIdColumnName: "Parent Issue",

    idDoneCalculator: () => {
        throw new Error('idDoneCalculator is not set')
    },

    stringFields: {},
    booleanFields: {},

    childIssueMetrics: [],
    blockerIssueMetrics: [],

    issueIdsExtractor: () => {
        throw new Error('issueIdsExtractor is not set')
    },
    issueIdDecorator: (id) => id,
    issueIdToUrl: () => {
        throw new Error('issueIdToUrl is not set')
    },

    issuesLoader: () => {
        throw new Error('issuesLoader is not set')
    },
    childIssuesLoader: () => {
        throw new Error('childIssuesLoader is not set')
    },
    blockerIssuesLoader: () => {
        throw new Error('blockerIssuesLoader is not set')
    },

    issueIdGetter: () => {
        throw new Error('issueIdGetter is not set')
    },

}

interface GSheetProjectSettings {

    firstDataRow: number

    settingsSheetName: string

    issueIdColumnName: string
    parentIssueIdColumnName: string

    isDoneColumnName?: string
    idDoneCalculator: IssueIsDoneCalculator

    stringFields: Record<string, IssueStringFieldGetter>
    booleanFields: Record<string, IssueBooleanFieldGetter>

    childIssueMetrics: IssueMetric[]
    blockerIssueMetrics: IssueMetric[]

    issueIdsExtractor: IssueIdsExtractor
    issueIdDecorator: IssueIdDecorator
    issueIdToUrl: IssueIdToUrl
    issueIdsToUrl?: IssueIdsToUrl

    issuesLoader: IssuesLoader
    childIssuesLoader: IssuesLoader
    blockerIssuesLoader: IssuesLoader

    issueIdGetter: IssueStringFieldGetter

}
