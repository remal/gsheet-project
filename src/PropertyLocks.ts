class PropertyLocks {

    static waitLock(property: string, timeout: number = GSheetProjectSettings.lockTimeoutMillis): boolean {
        property = `lock|${property}`

        const start = Date.now()
        while (true) {
            const propertyValue = PropertiesService.getDocumentProperties().getProperty(property)
            if (!propertyValue?.length) {
                break
            }
            const date = Utils.parseDate(propertyValue)
            if (date == null || date.getTime() < Date.now()) {
                break
            }

            if (start + timeout > Date.now()) {
                Utilities.sleep(1_000)
            } else {
                return false
            }
        }

        PropertiesService.getDocumentProperties().setProperty(property, Date.now().toString())
        return true
    }

    static releaseLock(property: string) {
        property = `lock|${property}`
        PropertiesService.getDocumentProperties().deleteProperty(property)
    }

    static releaseExpiredPropertyLocks() {
        for (const [key, value] of Object.entries(PropertiesService.getDocumentProperties().getProperties())) {
            if (!key.startsWith('lock|')) {
                continue
            }

            const date = Utils.parseDate(value)
            if (date == null || date.getTime() < Date.now()) {
                PropertiesService.getDocumentProperties().deleteProperty(key)
            }
        }
    }

}
