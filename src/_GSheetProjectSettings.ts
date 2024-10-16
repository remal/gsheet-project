class GSheetProjectSettings {

    static titleRow: Row = 1
    static firstDataRow: Row = this.titleRow + 1

    static lockColumns: boolean = false
    static lockRows: boolean = false
    static updateConditionalFormatRules: boolean = true
    static skipHiddenIssues: boolean = true
    static rewriteExistingDefaultFormula: boolean = false

    static issuesRangeName: RangeName = 'Issues'
    static childIssuesRangeName: RangeName = 'ChildIssues'
    static milestonesRangeName: RangeName = "Milestones"
    static titlesRangeName: RangeName = "Titles"
    static teamsRangeName: RangeName = "Teams"
    static estimatesRangeName: RangeName = "Estimates"
    static startsRangeName: RangeName = "Starts"
    static endsRangeName: RangeName = "Ends"
    static earliestStartsRangeName: RangeName = "EarliestStarts"
    static deadlinesRangeName: RangeName = "Deadlines"
    static warningDeadlinesRangeName: RangeName = "WarningDeadlines"

    static inProgressesRangeName: RangeName | undefined = undefined
    static codeCompletesRangeName: RangeName | undefined = undefined

    static settingsScheduleStartRangeName: RangeName = 'ScheduleStart'
    static settingsScheduleBufferRangeName: RangeName = 'ScheduleBuffer'
    static settingsScheduleWarningBufferRangeName: RangeName = 'ScheduleWarningBuffer'
    static settingsScheduleWarningBufferEstimateCoefficientRangeName: RangeName = 'ScheduleWarningBufferEstimateCoefficient'

    static settingsTeamsTableRangeName: RangeName = 'TeamsTable'
    static settingsTeamsTableTeamRangeName: RangeName = 'TeamsTableTeam'
    static settingsTeamsTableResourcesRangeName: RangeName = 'TeamsTableResources'

    static settingsMilestonesTableRangeName: RangeName = 'MilestonesTable'
    static settingsMilestonesTableMilestoneRangeName: RangeName = 'MilestonesTableMilestone'
    static settingsMilestonesTableDeadlineRangeName: RangeName = 'MilestonesTableDeadline'

    static publicHolidaysRangeName: RangeName = 'PublicHolidays'


    static notIssueKeyRegex: (RegExp | undefined) = /^\s*\W/
    static bufferIssueKeyRegex: (RegExp | undefined) = /^(buffer|reserve)/i
    static issueTrackers: IssueTracker[] = []
    static issuesLoadTimeoutMillis: number = 5 * 60 * 1000
    static onIssuesLoadedHandlers: OnIssuesLoaded[] = []
    static issuesMetrics: Record<ColumnName, IssuesMetric<any>> = {}
    static counterIssuesMetrics: Record<ColumnName, IssuesCounterMetric> = {}
    static originalIssueKeysTextChangedTimeout: number = 200

    static useLockService: boolean = true
    static lockTimeoutMillis: number = 5 * 60 * 1000


    static sheetName: SheetName = "Projects"
    static iconColumnName: ColumnName = "icon"
    //static doneColumnName: ColumnName = "Done"
    static milestoneColumnName: ColumnName = "Milestone"
    static typeColumnName: ColumnName = "Type"
    static issueKeyColumnName: ColumnName = "Issue"
    static childIssueKeyColumnName: ColumnName = "Child\nIssue"
    static lastDataReloadColumnName: ColumnName = "Last\nReload"
    static titleColumnName: ColumnName = "Title"
    static teamColumnName: ColumnName = "Team"
    static estimateColumnName: ColumnName = "Estimate\n(days)"
    static earliestStartColumnName: ColumnName = "Earliest\nStart"
    static deadlineColumnName: ColumnName = "Deadline"
    static warningDeadlineColumnName: ColumnName = "Warning\nDeadline"
    static startColumnName: ColumnName = "Start"
    static endColumnName: ColumnName = "End"
    //static issueHashColumnName: ColumnName = "Issue Hash"

    static settingsSheetName: SheetName = "Settings"

    static additionalColumns: ColumnInfo[] = []

    static daysTillDeadlineEstimateBufferDivider: number = 5

    static loadingText: string | undefined | null = '\u2B6E' // alternative: '\uD83D\uDD03'
    static indent: number = 4
    static fontSize: FontSize = 10

    // see https://spreadsheet.dev/how-to-get-the-hexadecimal-codes-of-colors-in-google-sheets
    static errorColor: Color = '#ff0000'
    static importantWarningColor: Color = '#e06666'
    static warningColor: Color = '#b45f06'
    static infoColor: Color = '#0000ff'
    static unimportantWarningColor: Color = '#fce5cd'
    static unimportantColor: Color = '#b7b7b7'


    static computeStringSettingsHash(): string {
        const hashableValues: Record<string, any> = {}
        const keys = Object.keys(GSheetProjectSettings).toSorted()
        for (const key of keys) {
            let value = GSheetProjectSettings[key]
            if (value instanceof RegExp) {
                value = value.toString()
            }

            if (value == null
                || typeof value === 'function'
                || typeof value === 'object'
            ) {
                continue
            }

            hashableValues[key] = value
        }

        const json = JSON.stringify(hashableValues)
        return SHA256(json)
    }

}
