const DEFAULT_GSHEET_PROJECT_SETTINGS: Partial<GSheetProjectSettings> = {

    firstDataRow: 2,

    settingsSheetName: "Settings",

    issueIdColumnName: "Issue",
    parentIssueIdColumnName: "Parent Issue",

    idDoneCalculator: () => Utils.throwNotConfigured("idDoneCalculator"),

    stringFields: {},
    booleanFields: {},

    childIssueMetrics: [],
    blockerIssueMetrics: [],

    issueIdsExtractor: () => Utils.throwNotConfigured("issueIdsExtractor"),
    issueIdDecorator: (id) => id,
    issueIdToUrl: () => Utils.throwNotConfigured("issueIdToUrl"),

    issuesLoader: () => Utils.throwNotConfigured("issuesLoader"),
    childIssuesLoader: () => Utils.throwNotConfigured("childIssuesLoader"),
    blockerIssuesLoader: () => Utils.throwNotConfigured("blockerIssuesLoader"),

    issueIdGetter: () => Utils.throwNotConfigured("issueIdGetter"),

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
