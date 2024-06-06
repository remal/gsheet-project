class SheetLayouts {

    private static readonly instances: ReadonlyArray<SheetLayout> = [
        SheetLayoutProjects.instance,
        SheetLayoutSettings.instance,
    ]

    static migrateColumnsIfNeeded() {
        this.instances.forEach(instance => instance.migrateColumnsIfNeeded())
    }

    static migrateColumns() {
        this.instances.forEach(instance => instance.migrateColumns())
    }

}
