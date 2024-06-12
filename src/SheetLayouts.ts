class SheetLayouts {

    private static get instances(): ReadonlyArray<SheetLayout> {
        return [
            SheetLayoutProjects.instance,
            SheetLayoutSettings.instance,
        ]
    }

    static migrateIfNeeded() {
        this.instances.forEach(instance => instance.migrateIfNeeded())
        CommonFormatter.applyCommonFormatsToAllSheets()
    }

    static migrate() {
        this.instances.forEach(instance => instance.migrate())
        CommonFormatter.applyCommonFormatsToAllSheets()
    }

}
