class GSheetProjectSettings {

    static titleRow: number = 1
    static firstDataRow: number = 2

    static sheetName: string = "Projects"
    static iconColumnName: string = "icon"
    static doneColumnName: string = "Done"
    static milestoneColumnName: string = "Milestone"
    static typeColumnName: string = "Type"
    static issueColumnName: string = "Issue"
    static issuesRangeName: string = "Issues"
    static childIssueColumnName: string = "Child\nIssue"
    static childIssuesRangeName: string = "ChildIssues"
    static titleColumnName: string = "Title"
    static teamColumnName: string = "Team"
    static teamsRangeName: string = "Teams"
    static estimateColumnName: string = "Estimate\n(days)"
    static estimatesRangeName: string = "Estimates"
    static deadlineColumnName: string = "Deadline"
    static deadlinesRangeName: string = "Deadlines"
    static startColumnName: string = "Start"
    static startsRangeName: string = "Starts"
    static endColumnName: string = "End"
    static endsRangeName: string = "Ends"
    //static issueHashColumnName: string = "Issue Hash"

    static indent: number = 4

    static taskTrackers: TaskTracker[] = []


    static computeStringSettingsHash(): string {
        const hashableValues: Record<string, any> = {}
        for (const [key, value] of Object.entries(GSheetProjectSettings)) {
            if (Utils.isString(value)) {
                hashableValues[key] = value
            }
        }

        const json = JSON.stringify(hashableValues)
        return SHA256(json)
    }

}
