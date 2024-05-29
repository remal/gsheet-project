class ProjectSheetLayout extends AbstractSheetLayout {

    static instance = new ProjectSheetLayout()

    protected get sheetName(): string {
        return GSheetProjectSettings.projectsSheetName
    }

    protected get columns(): ColumnInfo[] {
        return [
            {
                name: GSheetProjectSettings.projectsIssueColumnName,
                rangeName: GSheetProjectSettings.projectsIssueColumnRangeName,
            },
            {
                name: GSheetProjectSettings.projectsIssueHashColumnName,
                arrayFormula: '',
                rangeName: GSheetProjectSettings.projectsIssueHashColumnRangeName,
            },
        ]
    }

}
