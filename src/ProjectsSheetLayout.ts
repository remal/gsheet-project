class ProjectsSheetLayout extends SheetLayout {

    static readonly instance = new ProjectsSheetLayout()

    protected get sheetName(): string {
        return GSheetProjectSettings.projectsSheetName
    }

    protected get columns(): ReadonlyArray<ColumnInfo> {
        return [
            {
                name: GSheetProjectSettings.projectsIssueColumnName,
                rangeName: GSheetProjectSettings.projectsIssuesRangeName,
            },
            {
                name: GSheetProjectSettings.projectsIssueHashColumnName,
                arrayFormula: `
                    MAP(
                        ARRAYFORMULA(${GSheetProjectSettings.projectsIssuesRangeName}),
                        LAMBDA(issue, IF(ISBLANK(issue), "", ${SHA256.name}(issue)))
                    )
                `,
                rangeName: GSheetProjectSettings.projectsIssueHashesRangeName,
            },
        ]
    }

}
