class ScheduleSettings {

    static get start(): Date {
        const settings = Settings.getMap(GSheetProjectSettings.settingsScheduleScope)
        const startString = settings.get('start')
        if (!startString?.length) {
            return new Date()
        }

        return new Date(startString)
    }

}
