class GSheetProjectSettings {

    static titleRow: number = 1
    static firstDataRow: number = 2

    static settingsSheetName: string = "Settings"

    static projectsSheetName: string = "Projects"
    static projectsIconColumnName: string = "Icon"
    static projectsDoneColumnName: string = "Done"
    static projectsIssueColumnName: string = "Issue"
    static projectsIssuesRangeName: string = "Issues"
    static projectsChildIssueColumnName: string = "Child Issue"
    static projectsChildIssuesRangeName: string = "ChildIssues"
    static projectsTitleColumnName: string = "Title"
    static projectsTeamColumnName: string = "Team"
    static projectsEstimateColumnName: string = "Estimate (days)"
    static projectsDeadlineColumnName: string = "Deadline"
    static projectsStartColumnName: string = "Start"
    static projectsEndColumnName: string = "End"
    //static projectsIssueHashColumnName: string = "Issue Hash"

    static indent: number = 4
    static groupChildIssues: boolean = false

    static issueLoaderFactories: IssueLoaderFactory[] = []
    static issueChildrenLoaderFactories: IssueChildrenLoaderFactory[] = []
    static issueBlockersLoaderFactories: IssueBlockersLoaderFactory[] = []
    static issueSearcherFactories: IssueSearcherFactory[] = []


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
