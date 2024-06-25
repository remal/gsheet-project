class SheetLayouts {

    private static get instances(): ReadonlyArray<SheetLayout> {
        return [
            SheetLayoutProjects.instance,
            SheetLayoutSettings.instance,
        ]
    }

    private static _isMigrated: boolean = false

    static migrateIfNeeded() {
        Observability.timed([SheetLayouts.name, this.migrateIfNeeded.name].join(': '), () => {
            this.instances.forEach(instance => {
                const isMigrated = instance.migrateIfNeeded()
                if (isMigrated) {
                    this._isMigrated = true
                }
            })
            this.applyAfterMigrationSteps()
        })
    }

    static migrate() {
        if (this._isMigrated) {
            return
        }

        Observability.timed([SheetLayouts.name, this.migrate.name].join(': '), () => {
            this.instances.forEach(instance => instance.migrate())
            this.applyAfterMigrationSteps()
        })
    }

    static applyAfterMigrationSteps() {
        const rangeNames: RangeName[] = [
            GSheetProjectSettings.issuesRangeName,
            GSheetProjectSettings.childIssuesRangeName,
            GSheetProjectSettings.titlesRangeName,
            GSheetProjectSettings.teamsRangeName,
            GSheetProjectSettings.estimatesRangeName,
            GSheetProjectSettings.startsRangeName,
            GSheetProjectSettings.endsRangeName,
            GSheetProjectSettings.deadlinesRangeName,

            GSheetProjectSettings.inProgressesRangeName,
            GSheetProjectSettings.codeCompletesRangeName,

            GSheetProjectSettings.settingsScheduleStartRangeName,
            GSheetProjectSettings.settingsScheduleBufferRangeName,

            GSheetProjectSettings.settingsTeamsTableRangeName,
            GSheetProjectSettings.settingsTeamsTableTeamRangeName,
            GSheetProjectSettings.settingsTeamsTableResourcesRangeName,

            GSheetProjectSettings.settingsMilestonesTableRangeName,
            GSheetProjectSettings.settingsMilestonesTableMilestoneRangeName,
            GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName,

            GSheetProjectSettings.publicHolidaysRangeName,
        ].filter(it => it?.length).map(it => it!)
        const missingRangeNames = rangeNames.filter(name => NamedRangeUtils.findNamedRange(name) == null)
        if (missingRangeNames.length) {
            throw new Error(`Missing named range(s): '${missingRangeNames.join("', '")}'`)
        }

        CommonFormatter.applyCommonFormatsToAllSheets()
    }

}
