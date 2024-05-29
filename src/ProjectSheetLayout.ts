class ProjectSheetLayout extends AbstractSheetLayout {

    protected static get sheetName(): string {
        return GSheetProjectSettings.projectsSheetName
    }

    protected static get columns(): ColumnInfo[] {
        return [
            {
                name: GSheetProjectSettings.projectsIssueColumnName,
                rangeName: GSheetProjectSettings.projectsIssueColumnRangeName,
            },
        ]
    }

}
