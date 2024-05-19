class ScheduleSettings {

    static get start(): Date {
        const settings = Settings.getMap(GSheetProjectSettings.settingsScheduleScope)
        const stringValue = settings.get('start')
        if (!stringValue?.length) {
            return new Date()
        }

        return new Date(stringValue)
    }

    static get bufferCoefficient(): number {
        const settings = Settings.getMap(GSheetProjectSettings.settingsScheduleScope)
        const stringValue = settings.get('bufferCoefficient')
        const value = parseFloat(stringValue ?? '')
        return isNaN(value) ? 0 : value
    }

}
