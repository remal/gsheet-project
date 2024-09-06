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
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.childIssueKeyColumnName,
                rangeName: GSheetProjectSettings.childIssuesRangeName,
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(`=
                        #SELF_COLUMN(${GSheetProjectSettings.issuesRangeName})
                        =
                        OFFSET(#SELF_COLUMN(${GSheetProjectSettings.issuesRangeName}), -1, 0)
                    `)
                    .setHelpText(
                        `Children should be grouped under their parent`,
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
                conditionalFormats: [
                    builder => builder
                        .whenFormulaSatisfied(`=
                            OR(
                                NOT(ISNUMBER(#SELF)),
                                #SELF < 0
                            )
                        `)
                        .setFontColor(GSheetProjectSettings.unimportantColor),
                    builder => builder
                        .whenFormulaSatisfied(`=
                            AND(
                                #SELF = "",
                                #SELF_COLUMN(${GSheetProjectSettings.teamsRangeName}) <> ""
                            )
                        `)
                        .setBackground(GSheetProjectSettings.importantWarningColor),
                ],
            },
            {
                name: GSheetProjectSettings.startColumnName,
                rangeName: GSheetProjectSettings.startsRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
                conditionalFormats: [
                    GSheetProjectSettings.inProgressesRangeName?.length
                        ? builder => builder
                            .whenFormulaSatisfied(`=
                                AND(
                                    #SELF <> "",
                                    #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}) <> "",
                                    #SELF > #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}),
                                    #SELF_COLUMN(${GSheetProjectSettings.inProgressesRangeName}) <> "",
                                    ISFORMULA(#SELF),
                                    FORMULATEXT(#SELF) <> "=TODAY()"
                                )
                            `)
                            .setBold(true)
                            .setFontColor(GSheetProjectSettings.errorColor)
                            .setItalic(true)
                            .setBackground(GSheetProjectSettings.unimportantWarningColor)
                        : null,
                    builder => builder
                        .whenFormulaSatisfied(`=
                            AND(
                                #SELF <> "",
                                #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}) <> "",
                                #SELF > #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}),
                                ISFORMULA(#SELF),
                                FORMULATEXT(#SELF) <> "=TODAY()"
                            )
                        `)
                        .setBold(true)
                        .setFontColor(GSheetProjectSettings.errorColor)
                        .setItalic(true),
                    GSheetProjectSettings.inProgressesRangeName?.length
                        ? builder => builder
                            .whenFormulaSatisfied(`=
                                AND(
                                    #SELF_COLUMN(${GSheetProjectSettings.inProgressesRangeName}) <> "",
                                    ISFORMULA(#SELF),
                                    FORMULATEXT(#SELF) <> "=TODAY()",
                                    #SELF <> ""
                                )
                            `)
                            .setItalic(true)
                            .setBackground(GSheetProjectSettings.unimportantWarningColor)
                        : null,
                ],
            },
            {
                name: GSheetProjectSettings.endColumnName,
                rangeName: GSheetProjectSettings.endsRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
                conditionalFormats: [
                    builder => builder
                        .whenFormulaSatisfied(`=
                            AND(
                                #SELF <> "",
                                #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}) <> "",
                                #SELF > #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName})
                            )
                        `)
                        .setBold(true)
                        .setFontColor(GSheetProjectSettings.errorColor),
                    GSheetProjectSettings.codeCompletesRangeName?.length
                        ? builder => builder
                            .whenFormulaSatisfied(`=
                                AND(
                                    #SELF_COLUMN(${GSheetProjectSettings.codeCompletesRangeName}) = "",
                                    #SELF <> "",
                                    #SELF < TODAY()
                                )
                            `)
                            .setBold(true)
                            .setBackground(GSheetProjectSettings.unimportantWarningColor)
                            .setFontColor(GSheetProjectSettings.warningColor)
                        : null,
                ],
            },
            {
                name: GSheetProjectSettings.earliestStartColumnName,
                rangeName: GSheetProjectSettings.earliestStartsRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.deadlineColumnName,
                rangeName: GSheetProjectSettings.deadlinesRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.daysTillDeadlineColumnName,
                rangeName: GSheetProjectSettings.daysTillDeadlinesRangeName,
                hiddenByDefault: true,
                arrayFormula: `
                    ARRAYFORMULA(
                        IF(
                            (${GSheetProjectSettings.endsRangeName} <> "") * (${GSheetProjectSettings.deadlinesRangeName} <> ""),
                            NETWORKDAYS(${GSheetProjectSettings.endsRangeName}, ${GSheetProjectSettings.deadlinesRangeName}, ${GSheetProjectSettings.publicHolidaysRangeName}),
                            ""
                        )
                    )
                `,
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
            ...GSheetProjectSettings.additionalColumns,
        ]
    }

}
