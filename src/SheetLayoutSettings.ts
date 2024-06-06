class SheetLayoutSettings extends SheetLayout {

    static readonly instance = new SheetLayoutSettings()

    protected get sheetName(): string {
        return GSheetProjectSettings.settingsSheetName
    }

    protected get columns(): ReadonlyArray<ColumnInfo> {
        return []
    }

}
