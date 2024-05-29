class ProjectSheetLayout extends AbstractSheetLayout {

    static readonly instance = new ProjectSheetLayout()

    protected get sheetName(): string {
        return GSheetProjectSettings.projectsSheetName
    }

    protected get columns(): ColumnInfo[] {
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
                        LAMBDA(issue, IF(ISBLANK(issue), "", SHA256(issue)))
                    )
                `,
                rangeName: GSheetProjectSettings.projectsIssueHashesRangeName,
            },
        ]
    }

}
