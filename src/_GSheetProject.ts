function refreshSelectedRowsOfGSheetProject() {
    const range = SpreadsheetApp.getActiveRange()
    if (range == null) {
        return
    }

    const sheet = range.getSheet()
    if (!SheetUtils.isGridSheet(sheet)) {
        return
    }

    EntryPoint.entryPoint(() => {
        const rowsRange = sheet.getRange(
            `${range.getRow()}:${range.getRow() + range.getNumRows() - 1}`,
        )

        onEditGSheetProject({
            range: rowsRange,
        })
    })
}

function refreshAllRowsOfGSheetProject() {
    EntryPoint.entryPoint(() => {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
                const rowsRange = sheet.getRange(
                    `1:${SheetUtils.getLastRow(sheet)}`,
                )

                onEditGSheetProject({
                    range: rowsRange,
                })
            })
    }, false)
}

function reapplyDefaultFormulasOfGSheetProject() {
    EntryPoint.entryPoint(() => {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
                const rowsRange = sheet.getRange(
                    `1:${SheetUtils.getLastRow(sheet)}`,
                )

                DefaultFormulas.insertDefaultFormulas(rowsRange, true)
            })
    })
}

function applyDefaultStylesOfGSheetProject() {
    EntryPoint.entryPoint(() => {
        SheetLayouts.migrate()
    })
}

function cleanupGSheetProject() {
    EntryPoint.entryPoint(() => {
        SheetLayouts.migrateIfNeeded()
        ConditionalFormatting.removeDuplicateConditionalFormatRules()
        ConditionalFormatting.combineConditionalFormatRules()
        ProtectionLocks.releaseExpiredLocks()
        PropertyLocks.releaseExpiredPropertyLocks()
    }, false)
}

function onOpenGSheetProject(event?: SheetsOnOpen) {
    EntryPoint.entryPoint(() => {
        SheetLayouts.migrateIfNeeded()
    }, false)

    SpreadsheetApp.getUi()
        .createMenu("GSheetProject")
        .addItem("Refresh selected rows", refreshSelectedRowsOfGSheetProject.name)
        .addItem("Refresh all rows", refreshAllRowsOfGSheetProject.name)
        .addItem("Reapply default formulas", reapplyDefaultFormulasOfGSheetProject.name)
        .addItem("Apply default styles", applyDefaultStylesOfGSheetProject.name)
        .addToUi()
}

function onChangeGSheetProject(event?: SheetsOnChange) {
    function onInsert() {
        EntryPoint.entryPoint(() => {
            CommonFormatter.applyCommonFormatsToAllSheets()
        })
    }

    function onRemove() {
        applyDefaultStylesOfGSheetProject()
    }


    const changeType = event?.changeType?.toString() ?? ''
    if (['INSERT_ROW', 'INSERT_COLUMN'].includes(changeType)) {
        onInsert()
    } else if (['REMOVE_COLUMN'].includes(changeType)) {
        onRemove()
    }
}

function onEditGSheetProject(event?: Partial<Pick<SheetsOnEdit, 'range'>>) {
    const range = event?.range
    if (range == null) {
        return
    }

    EntryPoint.entryPoint(() => {
        Observability.timed(`Common format`, () => CommonFormatter.applyCommonFormatsToRowRange(range))
        //Observability.timed(`Done logic`, () => DoneLogic.executeDoneLogic(range))
        Observability.timed(`Default formulas`, () => DefaultFormulas.insertDefaultFormulas(range))
        Observability.timed(`Issue hierarchy`, () => IssueHierarchyFormatter.formatHierarchy(range))
        Observability.timed(`Reload issue data`, () => IssueDataDisplay.reloadIssueData(range))
    })
}

function onFormSubmitGSheetProject(event?: SheetsOnFormSubmit) {
    onEditGSheetProject({
        range: event?.range,
    })
}
