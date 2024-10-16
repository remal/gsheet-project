class DefaultFormulas extends AbstractIssueLogic {

    private static readonly _DEFAULT_FORMULA_MARKER = "default"
    private static readonly _DEFAULT_CHILD_FORMULA_MARKER = "default-child"
    private static readonly _DEFAULT_BUFFER_FORMULA_MARKER = "default-buffer"

    static isDefaultFormula(formula: string | null | undefined): boolean {
        return Formulas.extractFormulaMarkers(formula).includes(this._DEFAULT_FORMULA_MARKER)
    }

    static isDefaultChildFormula(formula: string | null | undefined): boolean {
        return Formulas.extractFormulaMarkers(formula).includes(this._DEFAULT_CHILD_FORMULA_MARKER)
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
        const endRow = range.getLastRow()

        const {issues, childIssues} = this._getIssueValues(sheet.getRange(
          GSheetProjectSettings.firstDataRow,
          range.getColumn(),
          Math.max(endRow - GSheetProjectSettings.firstDataRow + 1, 1),
          range.getNumColumns()
        ))

        const getParentIssueRow = (issueIndex: number): (number | undefined) => {
            const issue = issues[issueIndex]
            if (!issue?.length) {
                return undefined
            }

            let parentIssueIndex = issues.findLastIndex((curIssue, curIndex) =>
                curIndex < issueIndex
                && curIssue === issue
                && !childIssues[curIndex]?.length
            )
            if (parentIssueIndex < 0) {
                parentIssueIndex = issues.findIndex((curIssue, curIndex) =>
                    curIssue === issue
                    && !childIssues[curIndex]?.length
                )
            }
            if (parentIssueIndex < 0) {
                return undefined
            }

            return GSheetProjectSettings.firstDataRow + parentIssueIndex
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
        const earliestStartColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.earliestStartColumnName)
        const deadlineColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.deadlineColumnName)
        const warningDeadlineColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.warningDeadlineColumnName)


        const allValuesColumns = {} as Record<string, Column>
        ;[
            milestoneColumn,
            typeColumn,
            titleColumn,
            teamColumn,
            estimateColumn,
            startColumn,
            endColumn,
            earliestStartColumn,
            deadlineColumn,
            warningDeadlineColumn,
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
                const issueIndex = row - GSheetProjectSettings.firstDataRow
                const issue = issues[issueIndex]?.toString()
                const childIssue = childIssues[issueIndex]?.toString()

                const index = row - startRow
                let value = values[index]?.toString()
                let formula = formulas[index]


                if (GSheetProjectSettings.notIssueKeyRegex?.test(issue ?? '')
                    || (!issue?.length && !childIssue?.length)
                ) {
                    if (formula?.length) {
                        console.info([
                            DefaultFormulas.name,
                            sheet.getSheetName(),
                            addFormulas.name,
                            `column #${column}`,
                            `row #${row}`,
                            'cleaning formula for empty row',
                        ].join(': '))
                        sheet.getRange(row, column).setFormula('')
                    }
                    continue
                }


                const isChild = !!childIssue?.length
                const isDefaultFormula = this.isDefaultFormula(formula)
                const isDefaultChildFormula = this.isDefaultChildFormula(formula)
                if ((isChild && isDefaultFormula)
                    || (!isChild && isDefaultChildFormula)
                    || (rewriteExistingDefaultFormula && (isDefaultFormula || isDefaultChildFormula))
                ) {
                    value = ''
                    formula = ''
                }


                if (!value?.length && !formula?.length) {
                    const isBuffer = !!GSheetProjectSettings.bufferIssueKeyRegex?.test(issue ?? '')
                    console.info([
                        DefaultFormulas.name,
                        sheet.getSheetName(),
                        addFormulas.name,
                        `column #${column}`,
                        `row #${row}`,
                        `currentValue='${values[index]?.toString()}', isChild=${isChild}, isDefaultFormula=${isDefaultFormula}, isDefaultChildFormula=${isDefaultChildFormula}, isBuffer=${isBuffer}`,
                    ].join(': '))
                    let formula = Formulas.processFormula(formulaGenerator(
                        row,
                        isBuffer,
                        isChild,
                        issueIndex,
                        index,
                    ) ?? '')
                    if (formula.length) {
                        formula = Formulas.addFormulaMarkers(
                            formula,
                            isChild ? this._DEFAULT_CHILD_FORMULA_MARKER : this._DEFAULT_FORMULA_MARKER,
                            isBuffer ? this._DEFAULT_BUFFER_FORMULA_MARKER : null,
                        )
                        sheet.getRange(row, column).setFormula(formula)
                    } else {
                        sheet.getRange(row, column).setFormula('')
                    }
                }
            }
        }


        addFormulas(childIssueColumn, (row, isBuffer, isChild, issueIndex) => {
            if (isBuffer) {
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


        addFormulas(milestoneColumn, (row, isBuffer, isChild, issueIndex) => {
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

        addFormulas(typeColumn, (row, isBuffer, isChild, issueIndex) => {
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

        addFormulas(titleColumn, (row, isBuffer, isChild, issueIndex) => {
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

        addFormulas(estimateColumn, (row, isBuffer) => {
            if (isBuffer) {
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

        addFormulas(startColumn, (row, isBuffer) => {
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
            const earliestStartA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, earliestStartColumn))
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
                    MAX(${nextWorkdayLastEnd}, ${GSheetProjectSettings.settingsScheduleStartRangeName})
                )
            `

            let mainCalculation = `
                LET(
                    start,
                    ${firstDataRowIf},
                    IF(
                        ${earliestStartA1Notation} <> "",
                        MAX(start, ${earliestStartA1Notation}),
                        start
                    )
                )
            `

            if (isBuffer) {
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
                        ${firstDataRowIf}
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
                    IF(
                        resources,
                        ${mainCalculation},
                        ""
                    )
                )
            `

            const notEnoughDataIf = `
                IF(
                    ${teamA1Notation} = "",
                    "",
                    ${withResources}
                )
            `

            return `=${notEnoughDataIf}`
        })

        addFormulas(endColumn, (row, isBuffer) => {
            if (isBuffer) {
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
                        NOT(ISNUMBER(${estimateA1Notation})),
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

        addFormulas(deadlineColumn, (row, isBuffer, isChild, issueIndex) => {
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

        addFormulas(warningDeadlineColumn, (row, isBuffer, isChild, issueIndex) => {
            if (isBuffer) {
                return ''
            }

            const deadlineA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, deadlineColumn))
            const estimateA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, estimateColumn))
            const earliestStartA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, earliestStartColumn))
            return `=
                IF(
                    ${deadlineA1Notation} <> "",
                    LET(
                        warningBuffer,
                        ${GSheetProjectSettings.settingsScheduleWarningBufferRangeName} + IF(
                            N(${estimateA1Notation}) > 0,
                            FLOOR((${estimateA1Notation} - 1) / ${GSheetProjectSettings.settingsScheduleWarningBufferEstimateCoefficientRangeName}),
                            0
                        ),
                        IF(
                            OR(
                                ${earliestStartA1Notation} = "",
                                WORKDAY(
                                    ${earliestStartA1Notation},
                                    -1 * warningBuffer,
                                    ${GSheetProjectSettings.publicHolidaysRangeName}
                                ) <= WORKDAY(TODAY() + 1, -1, ${GSheetProjectSettings.publicHolidaysRangeName})
                            ),
                            WORKDAY(
                                ${deadlineA1Notation},
                                -1 * warningBuffer,
                                ${GSheetProjectSettings.publicHolidaysRangeName}
                            ),
                            ""
                        )
                    ),
                    ""
                )
            `
        })
    }

}

type DefaultFormulaGenerator = (
    row: Row,
    isBuffer: boolean,
    isChild: boolean,
    issueIndex: number,
    index: number,
) => string | null | undefined
