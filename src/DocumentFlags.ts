class DocumentFlags {

    static set(key: string, value: boolean = true) {
        if (value) {
            PropertiesService.getDocumentProperties().setProperty(key, new Date().getTime().toString())
        } else {
            PropertiesService.getDocumentProperties().deleteProperty(key)
        }
    }

    static isSet(key: string) {
        return PropertiesService.getDocumentProperties().getProperty(key)?.length
    }

    static cleanupByPrefix(keyPrefix: string) {
        interface Entry {
            key: string
            number: number
        }

        const entries: Entry[] = []
        for (const [key, value] of Object.entries(PropertiesService.getDocumentProperties().getProperties())) {
            if (key.startsWith(keyPrefix)) {
                const number = parseFloat(value)
                if (isNaN(number)) {
                    console.warn(`Removing NaN document flag: ${key}`)
                    PropertiesService.getDocumentProperties().deleteProperty(key)
                    continue
                }

                entries.push({key, number})
            }
        }

        // sort ascending:
        entries.sort((e1, e2) => e1.number - e2.number)

        // skip last element:
        entries.pop()

        // remove old keys:
        for (const entry of entries) {
            PropertiesService.getDocumentProperties().deleteProperty(entry.key)
        }
    }

}
