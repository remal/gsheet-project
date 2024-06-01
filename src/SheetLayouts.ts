class SheetLayouts {

    private static readonly instances: ReadonlyArray<SheetLayout> = [
        SheetLayoutProjects.instance,
    ]

    static migrateColumns() {
        this.instances.forEach(instance => instance.migrateColumns())
    }

}
