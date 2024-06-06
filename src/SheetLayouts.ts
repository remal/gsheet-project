class SheetLayouts {

    private static get instances(): ReadonlyArray<SheetLayout> {
        return [
            SheetLayoutProjects.instance,
            SheetLayoutSettings.instance,
        ]
    }

    static migrateColumnsIfNeeded() {
        this.instances.forEach(instance => instance.migrateColumnsIfNeeded())
    }

    static migrateColumns() {
        this.instances.forEach(instance => instance.migrateColumns())
    }

}
