class DefaultFormulas extends AbstractIssueLogic {

    private static readonly DEFAULT_FORMULA_MARKER = "default"

    static isDefaultFormula(formula: string | null | undefined): boolean {
        return Utils.extractFormulaMarkers(formula).includes(this.DEFAULT_FORMULA_MARKER)
    }

    static insertDefaultFormulas(range: Range) {
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

        const milestoneColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.milestoneColumnName)
        const teamColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.teamColumnName)
        const estimateColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName)
        const startColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.startColumnName)
        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName)
        const deadlineColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.deadlineColumnName)

        const addFormulas = (
            column: Column,
            formulaGenerator: (row: Row) => string,
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

                    if (!values[index]?.length && !formulas[index]?.length) {
                        console.info([
                            DefaultFormulas.name,
                            sheet.getSheetName(),
                            addFormulas.name,
                            `column #${column}`,
                            `row #${row}`,
                        ].join(': '))
                        let formula = Utils.processFormula(formulaGenerator(row))
                        formula = Utils.addFormulaMarker(formula, this.DEFAULT_FORMULA_MARKER)
                        sheet.getRange(row, column).setFormula(formula)
                    }
                }
            },
        )


        addFormulas(startColumn, row => {
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
                WORKDAY(${lastEnd}, 1, ${GSheetProjectSettings.publicHolidaysRangeName})
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

            const notEnoughDataIf = `
                IF(
                    ${teamA1Notation} = "",
                    "",
                    ${withResources}
                )
            `

            return `=${notEnoughDataIf}`
        })

        addFormulas(endColumn, row => {
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

        addFormulas(deadlineColumn, row => {
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
