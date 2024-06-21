class SheetLayoutProjects extends SheetLayout {

    static readonly instance = new SheetLayoutProjects()

    protected get sheetName(): string {
        return GSheetProjectSettings.sheetName
    }

    protected get columns(): ReadonlyArray<ColumnInfo> {
        return [
            {
                name: GSheetProjectSettings.iconColumnName,
                defaultTitleFontSize: 1,
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
                rangeName: GSheetProjectSettings.milestonesRangeName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.typeColumnName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.issueKeyColumnName,
                rangeName: GSheetProjectSettings.issuesRangeName,
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(`
                        =COUNTIFS(
                            ${GSheetProjectSettings.issuesRangeName}, "=" & #SELF,
                            ${GSheetProjectSettings.childIssuesRangeName}, "="
                        ) <= 1
                    `)
                    .setHelpText(
                        `Multiple rows with the same ${GSheetProjectSettings.issueKeyColumnName}`
                        + ` without ${GSheetProjectSettings.childIssueKeyColumnName}`,
                    )
                    .build(),
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.childIssueKeyColumnName,
                rangeName: GSheetProjectSettings.childIssuesRangeName,
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(`
                        =COUNTIF(${GSheetProjectSettings.issuesRangeName}, "=" & #SELF) = 0
                    `)
                    .setHelpText(
                        `Only one level of hierarchy is supported`,
                    )
                    .build(),
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.titleColumnName,
                rangeName: GSheetProjectSettings.titlesRangeName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.lastDataReloadColumnName,
                hiddenByDefault: true,
                defaultFormat: `yyyy-MM-dd HH:mm:ss.SSS`,
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.teamColumnName,
                rangeName: GSheetProjectSettings.teamsRangeName,
                //dataValidation <-- should be from ${GSheetProjectSettings.settingsTeamsTableTeamRangeName} range, see https://issuetracker.google.com/issues/143913035
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.estimateColumnName,
                rangeName: GSheetProjectSettings.estimatesRangeName,
                defaultFormat: '#,##0',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.startColumnName,
                rangeName: GSheetProjectSettings.startsRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.endColumnName,
                rangeName: GSheetProjectSettings.endsRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
                conditionalFormats: [
                    builder => builder
                        .whenFormulaSatisfied(
                            `=AND(
                                    ISFORMULA(#COLUMN_CELL),
                                    #COLUMN_CELL <> "",
                                    #COLUMN_CELL(deadline) <> "",
                                    #COLUMN_CELL > #COLUMN_CELL(deadline)
                                )`,
                        )
                        .setItalic(true)
                        .setBold(true)
                        .setFontColor('#c00'),
                    builder => builder
                        .whenFormulaSatisfied(
                            `=AND(
                                    #COLUMN_CELL <> "",
                                    #COLUMN_CELL(deadline) <> "",
                                    #COLUMN_CELL > #COLUMN_CELL(deadline)
                                )`,
                        )
                        .setBold(true)
                        .setFontColor('#f00'),
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
