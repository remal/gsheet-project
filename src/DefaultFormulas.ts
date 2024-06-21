class DefaultFormulas extends AbstractIssueLogic {

    private static readonly DEFAULT_FORMULA_MARKER = "default"
    private static readonly RESERVE_DEFAULT_FORMULA_MARKER = "reserve"

    static isDefaultFormula(formula: string | null | undefined): boolean {
        return Utils.extractFormulaMarkers(formula).includes(this.DEFAULT_FORMULA_MARKER)
    }

    static insertDefaultFormulas(range: Range, rewriteExistingDefaultFormula: boolean = false) {
        const processedRange = this._processRange(range)
        if (processedRange == null) {
            return
        } else {
            range = processedRange
        }

        const sheet = range.getSheet()
        const startRow = range.getRow()
        const endRow = startRow + range.getNumRows() - 1

        const {issues, childIssues} = this._getIssueValues(range)

        const childIssueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueKeyColumnName)
        const milestoneColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.milestoneColumnName)
        const teamColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.teamColumnName)
        const estimateColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName)
        const startColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.startColumnName)
        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName)
        const deadlineColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.deadlineColumnName)

        const addFormulas = (
            column: Column,
            formulaGenerator: (row: Row, isReserve: boolean) => string | null | undefined,
        ) => Observability.timed(
            [
                DefaultFormulas.name,
                sheet.getSheetName(),
                addFormulas.name,
                `column #${column}`,
            ].join(': '),
            () => {
                const values = this._getStringValues(range, column)
                const formulas = this._getFormulas(range, column)
                for (let row = startRow; row <= endRow; ++row) {
                    const index = row - startRow
                    if (!issues[index]?.length && !childIssues[index]?.length) {
                        if (formulas[index]?.length) {
                            sheet.getRange(row, column).setFormula('')
                        }
                        continue
                    }

                    if (rewriteExistingDefaultFormula && this.isDefaultFormula(formulas[index])) {
                        values[index] = ''
                        formulas[index] = ''
                    }

                    if (!values[index]?.length && !formulas[index]?.length) {
                        console.info([
                            DefaultFormulas.name,
                            sheet.getSheetName(),
                            addFormulas.name,
                            `column #${column}`,
                            `row #${row}`,
                        ].join(': '))

                        if (formulaGenerator != null) {
                            const isReserve = issues[index]?.startsWith(GSheetProjectSettings.reserveIssueKeyPrefix)
                            let formula = Utils.processFormula(formulaGenerator(row, isReserve) ?? '')
                            if (formula.length) {
                                formula = Utils.addFormulaMarker(formula, this.DEFAULT_FORMULA_MARKER)
                                sheet.getRange(row, column).setFormula(formula)
                            }
                        }
                    }
                }
            },
        )


        addFormulas(childIssueColumn, (row, isReserve) => {
            if (isReserve) {
                return `
                    =IF(
                        #SELF_COLUMN(${GSheetProjectSettings.teamsRangeName}) <> "",
                        #SELF_COLUMN(${GSheetProjectSettings.teamsRangeName})
                        & " - "
                        & COUNTIFS(
                            OFFSET(
                                ${GSheetProjectSettings.issuesRangeName},
                                0,
                                0,
                                ROW() - ${GSheetProjectSettings.firstDataRow} + 1,
                                1
                            ), "="&#SELF_COLUMN(${GSheetProjectSettings.issuesRangeName}),
                            OFFSET(
                                ${GSheetProjectSettings.teamsRangeName},
                                0,
                                0,
                                ROW() - ${GSheetProjectSettings.firstDataRow} + 1,
                                1
                            ), "="&#SELF_COLUMN(${GSheetProjectSettings.teamsRangeName})
                        ),
                        ""
                    )
                `
            }

            return undefined
        })

        addFormulas(estimateColumn, (row, isReserve) => {
            if (isReserve) {
                const startA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, startColumn))
                const endA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, endColumn))
                return `=LET(
                    workDays,
                    NETWORKDAYS(${startA1Notation}, ${endA1Notation}, ${GSheetProjectSettings.publicHolidaysRangeName}),
                    IF(
                        workDays > 0,
                        workDays,
                        ""
                    )
                )`
            }

            return undefined
        })

        addFormulas(startColumn, (row, isReserve) => {
            const teamTitleA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                GSheetProjectSettings.titleRow,
                teamColumn,
            ))
            const estimateTitleA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                GSheetProjectSettings.titleRow,
                estimateColumn,
            ))
            const endTitleA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                GSheetProjectSettings.titleRow,
                endColumn,
            ))

            const teamA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, teamColumn))
            const deadlineA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, deadlineColumn))

            const notEnoughPreviousLanes = `
                COUNTIFS(
                    OFFSET(
                        ${teamTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    "=" & ${teamA1Notation},
                    OFFSET(
                        ${estimateTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    ">0"
                ) < resources
            `

            const filter = `
                FILTER(
                    OFFSET(
                        ${endTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    OFFSET(
                        ${teamTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ) = ${teamA1Notation},
                    OFFSET(
                        ${estimateTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ) > 0
                )
            `

            const lastEnd = `
                MIN(
                    SORTN(
                        ${filter},
                        resources,
                        0,
                        1,
                        FALSE
                    )
                )
            `

            const nextWorkdayLastEnd = `
                WORKDAY(
                    ${lastEnd},
                    1,
                    ${GSheetProjectSettings.publicHolidaysRangeName}
                )
            `

            const firstDataRowIf = `
                IF(
                    OR(
                        ROW() <= ${GSheetProjectSettings.firstDataRow},
                        ${notEnoughPreviousLanes}
                    ),
                    ${GSheetProjectSettings.settingsScheduleStartRangeName},
                    ${nextWorkdayLastEnd}
                )
            `

            const withResources = `
                LET(
                    resources,
                    VLOOKUP(
                        ${teamA1Notation},
                        ${GSheetProjectSettings.settingsTeamsTableRangeName},
                        1
                            + COLUMN(${GSheetProjectSettings.settingsTeamsTableResourcesRangeName})
                            - COLUMN(${GSheetProjectSettings.settingsTeamsTableRangeName}),
                        FALSE
                    ),
                    ${firstDataRowIf}
                )
            `

            let mainCalculation = `
                LET(
                    dependencyEndDate,
                    DATE(2000, 1, 1),
                    MAX(
                        ${withResources},
                        WORKDAY(
                            dependencyEndDate,
                            1,
                            ${GSheetProjectSettings.publicHolidaysRangeName}
                        )
                    )
                )
            `

            if (isReserve) {
                let previousMilestone = `
                    MAX(FILTER(
                        ${GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName},
                        ${GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName} < ${deadlineA1Notation}
                    ))
                `
                previousMilestone = `
                    MAX(
                        ${previousMilestone},
                        ${GSheetProjectSettings.settingsScheduleStartRangeName} - 1
                    )
                `

                mainCalculation = `
                    MAX(
                        WORKDAY(
                            ${previousMilestone},
                            1,
                            ${GSheetProjectSettings.publicHolidaysRangeName}
                        ),
                        ${withResources}
                    )
                `

                mainCalculation = `
                    LET(
                        startDate,
                        ${mainCalculation},
                        IF(
                            startDate <= ${deadlineA1Notation},
                            startDate,
                            ""
                        )
                    )
                `
            }

            const notEnoughDataIf = `
                IF(
                    ${teamA1Notation} = "",
                    "",
                    ${mainCalculation}
                )
            `

            return `=${notEnoughDataIf}`
        })

        addFormulas(endColumn, (row, isReserve) => {
            if (isReserve) {
                const startA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, startColumn))
                const deadlineA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, deadlineColumn))
                return `=IF(
                    ${startA1Notation} <= ${deadlineA1Notation},
                    ${deadlineA1Notation},
                    ""
                )`
            }

            const startA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, startColumn))
            const estimateA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, estimateColumn))
            const bufferRangeName = GSheetProjectSettings.settingsScheduleBufferRangeName
            return `
                =IF(
                    OR(
                        ${startA1Notation} = "",
                        ${estimateA1Notation} = ""
                    ),
                    "",
                    WORKDAY(
                        ${startA1Notation},
                        MAX(ROUND(${estimateA1Notation} * (1 + ${bufferRangeName})) - 1, 0),
                        ${GSheetProjectSettings.publicHolidaysRangeName}
                    )
                )
            `
        })

        addFormulas(deadlineColumn, (row, isReserve) => {
            const milestoneA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, milestoneColumn))
            return `
                =IF(
                    ${milestoneA1Notation} = "",
                    "",
                    VLOOKUP(
                        ${milestoneA1Notation},
                        ${GSheetProjectSettings.settingsMilestonesTableRangeName},
                        1
                            + COLUMN(${GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName})
                            - COLUMN(${GSheetProjectSettings.settingsMilestonesTableRangeName}),
                        FALSE
                    )
                )
            `
        })
    }

}
