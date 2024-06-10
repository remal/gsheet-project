class DefaultFormulas extends AbstractIssueLogic {

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

        Utils.trimArrayEndBy(issues, it => !it?.length)
        SheetUtils.setLastRow(sheet, GSheetProjectSettings.firstDataRow + issues.length)
        range = RangeUtils.withMaxRow(range, GSheetProjectSettings.firstDataRow + issues.length)
        childIssues.length = issues.length

        const addFormulas = (
            column: number,
            formulaGenerator: (row: number) => string,
        ) => Utils.timed(
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
                        const formula = Utils.processFormula(formulaGenerator(row))
                        sheet.getRange(row, column).setFormula(formula)
                    }
                }
            },
        )


        const teamColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.teamColumnName)
        const estimateColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName)
        const startColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.startColumnName)
        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName)

        addFormulas(startColumn, row => {
            const teamFirstA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                GSheetProjectSettings.firstDataRow,
                teamColumn,
            ))
            const estimateFirstA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                GSheetProjectSettings.firstDataRow,
                estimateColumn,
            ))
            const endFirstA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(
                GSheetProjectSettings.firstDataRow,
                endColumn,
            ))

            const teamA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, teamColumn))
            const estimateA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, estimateColumn))

            const resourcesLookup = `
                VLOOKUP(
                    ${teamA1Notation},
                    ${GSheetProjectSettings.settingsTeamsTableRangeName},
                    1
                        + COLUMN(${GSheetProjectSettings.settingsTeamsTableResourcesRangeName})
                        - COLUMN(${GSheetProjectSettings.settingsTeamsTableRangeName}),
                    FALSE
                )
            `

            const notEnoughPreviousLanes = `
                COUNTIFS(
                    OFFSET(
                        ${teamFirstA1Notation},
                        0,
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    "=" & ${teamA1Notation},
                    OFFSET(
                        ${estimateFirstA1Notation},
                        0,
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    ">0"
                ) < ${resourcesLookup}
            `

            const filter = `
                FILTER(
                    OFFSET(
                        ${endFirstA1Notation},
                        0,
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    OFFSET(
                        ${teamFirstA1Notation},
                        0,
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ) = ${teamA1Notation},
                    OFFSET(
                        ${estimateFirstA1Notation},
                        0,
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
                        ${resourcesLookup},
                        0,
                        1,
                        FALSE
                    )
                )
            `

            const nextWorkdayOfQueryResult = `
                WORKDAY(
                    ${lastEnd},
                    1
                )
            `

            const firstDataRowIf = `
                IF(
                    OR(
                        ROW() <= ${GSheetProjectSettings.firstDataRow},
                        ${notEnoughPreviousLanes}
                    ),
                    ${GSheetProjectSettings.settingsScheduleStartRangeName},
                    ${nextWorkdayOfQueryResult}
                )
            `

            const notEnoughDataIf = `
                IF(
                    OR(
                        ${teamA1Notation} = "",
                        ${estimateA1Notation} = ""
                    ),
                    "",
                    ${firstDataRowIf}
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
                    WORKDAY(${startA1Notation}, ${estimateA1Notation} * (1 + ${bufferRangeName}))
                )
            `
        })
    }

}
