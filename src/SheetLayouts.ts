class SheetLayouts {

    private static get instances(): ReadonlyArray<SheetLayout> {
        return [
            SheetLayoutProjects.instance,
            SheetLayoutSettings.instance,
        ]
    }

    static migrateIfNeeded() {
        this.instances.forEach(instance => instance.migrateIfNeeded())
        this.applyAfterMigrationSteps()
    }

    static migrate() {
        this.instances.forEach(instance => instance.migrate())
        this.applyAfterMigrationSteps()
    }

    static applyAfterMigrationSteps() {
        const rangeNames = [
            GSheetProjectSettings.issuesRangeName,
            GSheetProjectSettings.childIssuesRangeName,
            GSheetProjectSettings.titlesRangeName,
            GSheetProjectSettings.teamsRangeName,
            GSheetProjectSettings.estimatesRangeName,
            GSheetProjectSettings.startsRangeName,
            GSheetProjectSettings.endsRangeName,

            GSheetProjectSettings.settingsScheduleStartRangeName,
            GSheetProjectSettings.settingsScheduleBufferRangeName,

            GSheetProjectSettings.settingsTeamsTableRangeName,
            GSheetProjectSettings.settingsTeamsTableTeamRangeName,
            GSheetProjectSettings.settingsTeamsTableResourcesRangeName,

            GSheetProjectSettings.settingsMilestonesTableRangeName,
            GSheetProjectSettings.settingsMilestonesTableMilestoneRangeName,
            GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName,

            GSheetProjectSettings.publicHolidaysRangeName,
        ]
        const missingRangeNames = rangeNames.filter(name => NamedRangeUtils.findNamedRange(name) == null)
        if (missingRangeNames.length) {
            throw new Error(`Missing named range(s): '${missingRangeNames.join("', '")}'`)
        }

        CommonFormatter.applyCommonFormatsToAllSheets()
    }

}
