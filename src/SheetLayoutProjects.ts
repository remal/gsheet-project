class SheetLayoutProjects extends SheetLayout {

    static readonly instance = new SheetLayoutProjects()

    protected get sheetName(): string {
        return GSheetProjectSettings.projectsSheetName
    }

    protected get columns(): ReadonlyArray<ColumnInfo> {
        return [
            {
                name: GSheetProjectSettings.projectsIconColumnName,
                defaultFontSize: 1,
                defaultWidth: '#default-height',
            },
            {
                name: GSheetProjectSettings.projectsDoneColumnName,
            },
            {
                name: GSheetProjectSettings.projectsIssueColumnName,
                rangeName: GSheetProjectSettings.projectsIssuesRangeName,
                dataValidation: () => SpreadsheetApp.newDataValidation().requireFormulaSatisfied(
                    `=OR(
                        NOT(ISBLANK(${GSheetProjectSettings.projectsChildIssuesRangeName})),
                        COUNTIFS(
                            ${GSheetProjectSettings.projectsIssuesRangeName}, "=" & #SELF,
                            ${GSheetProjectSettings.projectsChildIssuesRangeName}, "="
                        ) <= 1
                    )`,
                ).build(),
            },
            {
                name: GSheetProjectSettings.projectsChildIssueColumnName,
                rangeName: GSheetProjectSettings.projectsChildIssuesRangeName,
            },
            {
                name: GSheetProjectSettings.projectsTitleColumnName,
            },
            {
                name: GSheetProjectSettings.projectsTeamColumnName,
            },
            {
                name: GSheetProjectSettings.projectsEstimateColumnName,
            },
            {
                name: GSheetProjectSettings.projectsDeadlineColumnName,
            },
            {
                name: GSheetProjectSettings.projectsStartColumnName,
            },
            {
                name: GSheetProjectSettings.projectsEndColumnName,
            },
            /*
            {
                name: GSheetProjectSettings.projectsIssueHashColumnName,
                arrayFormula: `
                    MAP(
                        ARRAYFORMULA(${GSheetProjectSettings.projectsIssuesRangeName}),
                        LAMBDA(issue, IF(ISBLANK(issue), "", ${SHA256.name}(issue)))
                    )
                `,
            },
            */
        ]
    }

}
