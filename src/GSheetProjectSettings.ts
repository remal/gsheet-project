class GSheetProjectSettings {

    static titleRow: number = 1
    static firstDataRow: number = 2

    static settingsSheetName: string = "Settings"

    static projectsSheetName: string = "Projects"
    static projectsIssueColumnName: string = "Issue"
    static projectsIssueColumnRangeName: string = "Issues"


    static computeSettingsHash() {
        const hashableValues: Record<string, any> = {}
        for (const [key, value] of Object.entries(GSheetProjectSettings)) {
            if (value == null
                || typeof value === 'function'
                || typeof value === 'object'
            ) {
                continue
            }

            hashableValues[key] = value
        }

        const json = JSON.stringify(hashableValues)
        return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, json)
    }

}
