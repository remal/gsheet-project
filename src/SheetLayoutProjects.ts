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
            /*
            {
                name: GSheetProjectSettings.doneColumnName,
                defaultHorizontalAlignment: 'center',
            },
            */
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
                            ${GSheetProjectSettings.childIssuesRangeName} <> "",
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
                rangeName: GSheetProjectSettings.teamsRangeName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.estimateColumnName,
                rangeName: GSheetProjectSettings.estimatesRangeName,
                defaultFormat: '#,##0',
                defaultHorizontalAlignment: 'center',
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(
                        `=INDIRECT(ADDRESS(ROW(), COLUMN(${GSheetProjectSettings.teamsRangeName}))) <> ""`,
                    )
                    .setHelpText(
                        `Estimate must be defined for a team`,
                    )
                    .build(),
            },
            {
                name: GSheetProjectSettings.startColumnName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.endColumnName,
                rangeName: GSheetProjectSettings.endsRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
                conditionalFormats: [
                    {
                        order: 1,
                        configurer: builder => builder
                            .whenFormulaSatisfied(
                                `=AND(ISFORMULA(#COLUMN_CELL), NOT(#COLUMN_CELL = ""), #COLUMN_CELL > #COLUMN_CELL(deadline))`,
                            )
                            .setItalic(true)
                            .setBold(true)
                            .setFontColor('red'),
                    },
                    {
                        order: 2,
                        configurer: builder => builder
                            .whenFormulaSatisfied(
                                `=AND(NOT(#COLUMN_CELL = ""), #COLUMN_CELL > #COLUMN_CELL(deadline))`,
                            )
                            .setBold(true)
                            .setFontColor('red'),
                    },
                ],
            },
            {
                key: 'deadline',
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
                        LAMBDA(issue, IF(issue = "", "", ${SHA256.name}(issue)))
                    )
                `,
            },
            */
        ]
    }

}
