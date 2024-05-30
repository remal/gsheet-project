class SheetLayouts {

    private static readonly instances: ReadonlyArray<SheetLayout> = [
        ProjectsSheetLayout.instance,
    ]

    static migrateColumns() {
        this.instances.forEach(instance => instance.migrateColumns())
    }

}
