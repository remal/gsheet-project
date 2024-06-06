class SheetLayoutProjects extends SheetLayout {

    static readonly instance = new SheetLayoutProjects()

    protected get sheetName(): string {
        return GSheetProjectSettings.sheetName
    }

    protected get columns(): ReadonlyArray<ColumnInfo> {
        return [
            {
                name: GSheetProjectSettings.iconColumnName,
                defaultFontSize: 1,
                defaultWidth: '#default-height',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.doneColumnName,
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.milestoneColumnName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.typeColumnName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.issueColumnName,
                rangeName: GSheetProjectSettings.issuesRangeName,
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(
                        `=OR(
                            NOT(ISBLANK(${GSheetProjectSettings.childIssuesRangeName})),
                            COUNTIFS(
                                ${GSheetProjectSettings.issuesRangeName}, "=" & #SELF,
                                ${GSheetProjectSettings.childIssuesRangeName}, "="
                            ) <= 1
                        )`,
                    )
                    .setHelpText(
                        `Multiple rows with the same ${GSheetProjectSettings.issueColumnName}`
                        + ` without ${GSheetProjectSettings.childIssueColumnName}`,
                    )
                    .build(),
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.childIssueColumnName,
                rangeName: GSheetProjectSettings.childIssuesRangeName,
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(
                        `=COUNTIF(${GSheetProjectSettings.issuesRangeName}, "=" & #SELF) = 0`,
                    )
                    .setHelpText(
                        `Only one level of hierarchy is supported`,
                    )
                    .build(),
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.titleColumnName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.teamColumnName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.estimateColumnName,
                defaultFormat: '#,##0',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.startColumnName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.endColumnName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.deadlineColumnName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            /*
            {
                name: GSheetProjectSettings.projectsIssueHashColumnName,
                hiddenByDefault: true,
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
