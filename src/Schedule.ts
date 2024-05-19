class Schedule {

    static recalculateSchedule(range: Range) {
        if (!RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.estimateColumnName)
            || !RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.startColumnName)
            || !RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.endColumnName)
        ) {
            return
        }

        this._recalculateSheetSchedule(range.getSheet())
    }

    static recalculateAllSchedules() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            this._recalculateSheetSchedule(sheet)
        }
    }

    private static _recalculateSheetSchedule(sheet: Sheet) {
        if (State.isStructureChanged()) return
        const estimateColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.estimateColumnName)
        if (estimateColumn == null) {
            return
        }

        let lanesRange: Range | null = null
        const laneColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.laneColumnName)
        if (laneColumn != null) {
            lanesRange = SheetUtils.getColumnRange(
                sheet,
                laneColumn,
                GSheetProjectSettings.firstDataRow,
            ).setValue('')
        }

        const startsRange = SheetUtils.getColumnRange(
            sheet,
            GSheetProjectSettings.startColumnName,
            GSheetProjectSettings.firstDataRow,
        ).setValue('')
        const endsRange = SheetUtils.getColumnRange(
            sheet,
            GSheetProjectSettings.endColumnName,
            GSheetProjectSettings.firstDataRow,
        ).setValue('')

        const generalEstimatesRange = SheetUtils.getColumnRange(
            sheet,
            estimateColumn,
            GSheetProjectSettings.firstDataRow,
        )
        const generalEstimates: string[] = generalEstimatesRange.getValues()
            .map(cols => cols[0].toString().trim())

        const allDaysEstimates: DayEstimate[] = []
        const allTeamDaysEstimates = new Map<string, DayEstimate[]>
        const invalidEstimateRows: number[] = []
        generalEstimates.forEach((generalEstimate, index) => {
            if (!generalEstimate.length) {
                return
            }

            let isTeamFound = false
            for (const team of Team.getAll()) {
                const regex = new RegExp(`^${Utils.escapeRegex(team.id)}\\s*:\\s*(\\d+)\\s*([dw])?$`, 'i')
                const match = generalEstimate.match(regex)
                if (match == null) {
                    continue
                }

                const amount = parseInt(match[1])
                const unit = match[2]?.toUpperCase() ?? 'D'

                let days: number
                if (unit === 'W') {
                    days = amount * 5
                } else {
                    days = amount
                }

                const row = GSheetProjectSettings.firstDataRow + index

                const canonicalGeneralEstimate = `${team.id}: ${amount}${unit !== 'D' ? unit : ''}`
                if (canonicalGeneralEstimate !== generalEstimate) {
                    if (State.isStructureChanged()) return
                    sheet.getRange(row, estimateColumn).setValue(canonicalGeneralEstimate)
                }

                let teamDayEstimates = allTeamDaysEstimates.get(team.id)
                if (teamDayEstimates == null) {
                    teamDayEstimates = []
                    allTeamDaysEstimates.set(team.id, teamDayEstimates)
                }

                const dayEstimate: DayEstimate = {
                    index: index,
                    row: row,
                    daysEstimate: days,
                    teamId: team.id,
                }
                allDaysEstimates.push(dayEstimate)
                teamDayEstimates.push(dayEstimate)
                isTeamFound = true
                break
            }

            if (!isTeamFound) {
                invalidEstimateRows.push(GSheetProjectSettings.firstDataRow + index)
            }
        })

        if (State.isStructureChanged()) return
        generalEstimatesRange.setBackground(null)
        const invalidEstimateNotations = invalidEstimateRows
            .map(row => sheet.getRange(row, estimateColumn).getA1Notation())
        if (invalidEstimateNotations.length) {
            sheet.getRangeList(invalidEstimateNotations).setBackground('#FFCCCB')
        }

        const skipWeekends = (date: Date): Date => {
            while (date.getDay() === 0 || date.getDay() === 6) {
                date = new Date(date.getTime() + 24 * 3600 * 1000)
            }
            return date
        }

        const scheduleStart = skipWeekends(ScheduleSettings.start)
        const startsRangeValues = Utils.arrayOf<(Date | string)[]>(startsRange.getHeight(), [''])
        const endsRangeValues = Utils.arrayOf<(Date | string)[]>(endsRange.getHeight(), [''])
        for (const [teamId, teamDayEstimates] of allTeamDaysEstimates.entries()) {
            const lanes = new Lanes<DayEstimate>(Team.getById(teamId).lanes)
            teamDayEstimates.forEach(dayEstimate => lanes.add(
                dayEstimate.daysEstimate,
                dayEstimate,
                laneIndex => dayEstimate.laneIndex = laneIndex,
            ))

            for (const lane of lanes.lanes) {
                let lastEnd: Date | undefined = undefined
                for (const dayEstimate of lane.objects()) {
                    let start = scheduleStart
                    if (lastEnd != null) {
                        start = skipWeekends(new Date(lastEnd.getTime() + 24 * 3600 * 1000))
                    }
                    startsRangeValues[dayEstimate.index] = [start]

                    let end = start
                    const daysEstimate = Math.ceil(dayEstimate.daysEstimate * (1 + ScheduleSettings.bufferCoefficient))
                    for (let day = 1; day <= daysEstimate; ++day) {
                        end = skipWeekends(new Date(end.getTime() + 24 * 3600 * 1000))
                    }
                    endsRangeValues[dayEstimate.index] = [end]

                    lastEnd = end
                }
            }
        }
        if (State.isStructureChanged()) return
        startsRange.setValues(startsRangeValues)
        endsRange.setValues(endsRangeValues)

        if (lanesRange != null) {
            const laneRangeValues = Utils.arrayOf<string[]>(lanesRange.getHeight(), [''])
            for (const y of Utils.range(1, lanesRange.getHeight())) {
                const index = y - 1
                const dayEstimate = allDaysEstimates.find(it => it.index === index)
                if (dayEstimate?.laneIndex != null) {
                    laneRangeValues[index] = [`${dayEstimate.teamId}-${dayEstimate.laneIndex + 1}`]
                }
            }
            lanesRange.setValues(laneRangeValues)
        }
    }

}

interface DayEstimate {
    index: number
    row: number
    daysEstimate: number
    teamId: string
    laneIndex?: number
}

