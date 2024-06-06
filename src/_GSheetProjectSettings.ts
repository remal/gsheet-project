class GSheetProjectSettings {

    static titleRow: number = 1
    static firstDataRow: number = 2
    static restoreUndoneEnd: boolean = false

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
    static estimateColumnName: string = "Estimate\n(days)"
    static deadlineColumnName: string = "Deadline"
    static startColumnName: string = "Start"
    static endColumnName: string = "End"
    //static issueHashColumnName: string = "Issue Hash"

    static settingsSheetName: string = "Settings"

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
