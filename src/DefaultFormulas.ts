class DefaultFormulas extends AbstractIssueLogic {

    private static readonly DEFAULT_FORMULA_MARKER = "default"
    private static readonly DEFAULT_CHILD_FORMULA_MARKER = "default-child"

    static isDefaultFormula(formula: string | null | undefined): boolean {
        return Formulas.extractFormulaMarkers(formula).includes(this.DEFAULT_FORMULA_MARKER)
    }

    static isDefaultChildFormula(formula: string | null | undefined): boolean {
        return Formulas.extractFormulaMarkers(formula).includes(this.DEFAULT_CHILD_FORMULA_MARKER)
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
        const rows = range.getNumRows()
        const endRow = startRow + rows - 1

        const {issues, childIssues} = this._getIssueValues(sheet.getRange(
            GSheetProjectSettings.firstDataRow,
            range.getColumn(),
            endRow - GSheetProjectSettings.firstDataRow + 1,
            range.getNumColumns(),
        ))

        const getParentIssueRow = (issueIndex: number): (number | undefined) => {
            const issue = issues[issueIndex]
            if (!issue?.length) {
                return undefined
            }

            const index = issues.indexOf(issue)
            if (index < 0) {
                return undefined
            }

            const childIssue = childIssues[index]
            if (childIssue?.length) {
                return undefined
            }

            return GSheetProjectSettings.firstDataRow + index
        }

        const issueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName)
        const childIssueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueKeyColumnName)
        const childIssueFormulas = LazyProxy.create(() =>
            SheetUtils.getColumnsFormulas(sheet, {childIssues: childIssueColumn}, startRow, endRow).childIssues,
        )

        const milestoneColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.milestoneColumnName)
        const typeColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.typeColumnName)
        const titleColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.titleColumnName)
        const teamColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.teamColumnName)
        const estimateColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName)
        const startColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.startColumnName)
        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName)
        const deadlineColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.deadlineColumnName)


        const allValuesColumns = {} as Record<string, Column>
        ;[
            milestoneColumn,
            typeColumn,
            titleColumn,
            teamColumn,
            estimateColumn,
            startColumn,
            endColumn,
            deadlineColumn,
        ].forEach(column => allValuesColumns[column.toString()] = column)
        const allValues = LazyProxy.create(() =>
            SheetUtils.getColumnsStringValues(sheet, allValuesColumns, startRow, endRow),
        )
        const getValues = (column: Column): (string[]) => {
            if (column === issueColumn) {
                return issues.slice(-rows)
            } else if (column === childIssueColumn) {
                return childIssues.slice(-rows)
            }

            return allValues[column.toString()] ?? (() => {
                throw new Error(`Column ${column} is not prefetched`)
            })()
        }

        const allFormulas = LazyProxy.create(() =>
            SheetUtils.getColumnsFormulas(sheet, allValuesColumns, startRow, endRow),
        )
        const getFormulas = (column: Column): (string[]) => {
            if (column === childIssueColumn) {
                return childIssueFormulas
            }

            return allFormulas[column.toString()] ?? (() => {
                throw new Error(`Column ${column} is not prefetched`)
            })()
        }

        const addFormulas = (
            column: Column,
            formulaGenerator: DefaultFormulaGenerator,
        ) => {
            const values = getValues(column)
            const formulas = getFormulas(column)
            for (let row = startRow; row <= endRow; ++row) {
                const index = row - startRow
                const issueIndex = row - GSheetProjectSettings.firstDataRow
                if (!issues[issueIndex]?.length && !childIssues[issueIndex]?.length) {
                    if (formulas[issueIndex]?.length) {
                        sheet.getRange(row, column).setFormula('')
                    }
                    continue
                }

                const isChild = !!childIssues[index]?.length
                const isDefaultFormula = this.isDefaultFormula(formulas[index])
                const isDefaultChildFormula = this.isDefaultChildFormula(formulas[index])
                if ((isChild && isDefaultFormula)
                    || (!isChild && isDefaultChildFormula)
                    || (rewriteExistingDefaultFormula && (isDefaultFormula || isDefaultChildFormula))
                ) {
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
                    const isReserve = issues[index]?.startsWith(GSheetProjectSettings.reserveIssueKeyPrefix)
                    let formula = Formulas.processFormula(formulaGenerator(
                        row,
                        isReserve,
                        isChild,
                        issueIndex,
                        index,
                    ) ?? '')
                    if (formula.length) {
                        formula = Formulas.addFormulaMarker(
                            formula,
                            isChild ? this.DEFAULT_CHILD_FORMULA_MARKER : this.DEFAULT_FORMULA_MARKER,
                        )
                        sheet.getRange(row, column).setFormula(formula)
                    } else {
                        sheet.getRange(row, column).setFormula('')
                    }
                }
            }
        }


        addFormulas(milestoneColumn, (row, isReserve, isChild, issueIndex) => {
            if (isChild) {
                const parentIssueRow = getParentIssueRow(issueIndex)
                if (parentIssueRow != null) {
                    const milestoneA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                        parentIssueRow,
                        milestoneColumn,
                    ))
                    return `=${milestoneA1Notation}`
                }
            }

            return undefined
        })

        addFormulas(typeColumn, (row, isReserve, isChild, issueIndex) => {
            if (isChild) {
                const parentIssueRow = getParentIssueRow(issueIndex)
                if (parentIssueRow != null) {
                    const typeA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                        parentIssueRow,
                        typeColumn,
                    ))
                    return `=${typeA1Notation}`
                }
            }

            return undefined
        })

        addFormulas(childIssueColumn, (row, isReserve, isChild, issueIndex) => {
            if (isReserve) {
                childIssues[issueIndex] = `placeholder: ${addFormulas.name}`
                return `=
                    IF(
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

        addFormulas(titleColumn, (row, isReserve, isChild, issueIndex) => {
            if (isChild) {
                const parentIssueRow = getParentIssueRow(issueIndex)
                if (parentIssueRow != null) {
                    const titleA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                        parentIssueRow,
                        titleColumn,
                    ))
                    const childIssueA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                        row,
                        childIssueColumn,
                    ))
                    return `=${titleA1Notation} & " - " & ${childIssueA1Notation}`
                }

                const childIssueA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                    row,
                    childIssueColumn,
                ))
                return `=${childIssueA1Notation}`
            }

            const issueA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                row,
                issueColumn,
            ))
            return `=${issueA1Notation}`
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
            return `=
                IF(
                    OR(
                        ${startA1Notation} = "",
                        ${estimateA1Notation} = "",
                        ${estimateA1Notation} <= 0
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

        addFormulas(deadlineColumn, (row, isReserve, isChild, issueIndex) => {
            if (isChild) {
                const parentIssueRow = getParentIssueRow(issueIndex)
                if (parentIssueRow != null) {
                    const deadlineA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                        parentIssueRow,
                        deadlineColumn,
                    ))
                    return `=${deadlineA1Notation}`
                }
            }

            const milestoneA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, milestoneColumn))
            return `=
                IF(
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

type DefaultFormulaGenerator = (
    row: Row,
    isReserve: boolean,
    isChild: boolean,
    issueIndex: number,
    index: number,
) => string | null | undefined
