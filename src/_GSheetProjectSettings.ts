class GSheetProjectSettings {

    static titleRow: Row = 1
    static firstDataRow: Row = this.titleRow + 1

    static lockColumns: boolean = false
    static lockRows: boolean = true
    static updateConditionalFormatRules: boolean = true
    static reorderHierarchyAutomatically: boolean = false
    static skipHiddenIssues: boolean = true
    //static restoreUndoneEnd: boolean = false

    static issuesRangeName: RangeName = 'Issues'
    static childIssuesRangeName: RangeName = 'ChildIssues'
    static titlesRangeName: RangeName = "Titles"
    static teamsRangeName: RangeName = "Teams"
    static estimatesRangeName: RangeName = "Estimates"
    static startsRangeName: RangeName = "Starts"
    static endsRangeName: RangeName = "Ends"

    static settingsScheduleStartRangeName: RangeName = 'ScheduleStart'
    static settingsScheduleBufferRangeName: RangeName = 'ScheduleBuffer'

    static settingsTeamsTableRangeName: RangeName = 'TeamsTable'
    static settingsTeamsTableTeamRangeName: RangeName = 'TeamsTableTeam'
    static settingsTeamsTableResourcesRangeName: RangeName = 'TeamsTableResources'

    static settingsMilestonesTableRangeName: RangeName = 'MilestonesTable'
    static settingsMilestonesTableMilestoneRangeName: RangeName = 'MilestonesTableMilestone'
    static settingsMilestonesTableDeadlineRangeName: RangeName = 'MilestonesTableDeadline'

    static publicHolidaysRangeName: RangeName = 'PublicHolidays'


    static issueTrackers: IssueTracker[] = []
    static issuesLoadTimeoutMillis: number = 5 * 60 * 1000
    static booleanIssuesMetrics: Record<ColumnName, IssuesMetric<boolean>> = {}
    static counterIssuesMetrics: Record<ColumnName, IssuesCounterMetric> = {}

    static useLockService: boolean = true
    static lockTimeoutMillis: number = 10 * 60 * 1000


    static sheetName: SheetName = "Projects"
    static iconColumnName: ColumnName = "icon"
    //static doneColumnName: ColumnName = "Done"
    static milestoneColumnName: ColumnName = "Milestone"
    static typeColumnName: ColumnName = "Type"
    static issueColumnName: ColumnName = "Issue"
    static childIssueColumnName: ColumnName = "Child\nIssue"
    static lastDataReloadColumnName: ColumnName = "Last\nReload"
    static titleColumnName: ColumnName = "Title"
    static teamColumnName: ColumnName = "Team"
    static estimateColumnName: ColumnName = "Estimate\n(days)"
    static deadlineColumnName: ColumnName = "Deadline"
    static startColumnName: ColumnName = "Start"
    static endColumnName: ColumnName = "End"
    //static issueHashColumnName: ColumnName = "Issue Hash"

    static settingsSheetName: SheetName = "Settings"

    static loadingText: string | undefined | null = '\u2B6E' // alternative: '\uD83D\uDD03'
    static indent: number = 4
    static fontSize: FontSize = 10


    static computeStringSettingsHash(): string {
        const hashableValues: Record<string, any> = {}
        for (const [key, value] of Object.entries(GSheetProjectSettings)) {
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
