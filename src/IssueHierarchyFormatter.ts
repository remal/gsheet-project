class IssueHierarchyFormatter {

    static readonly FORMULA_MARKER = "hierarchy"

    static formatHierarchy(range: Range) {
        if (![GSheetProjectSettings.childIssueColumnName].some(columnName =>
            RangeUtils.doesRangeHaveSheetColumn(range, GSheetProjectSettings.sheetName, columnName),
        )) {
            return
        }

        let issuesRange = RangeUtils.toColumnRange(range, GSheetProjectSettings.issueColumnName)
        if (issuesRange == null) {
            return
        }

        const sheet = issuesRange.getSheet()
        issuesRange = RangeUtils.withMinMaxRows(
            issuesRange,
            GSheetProjectSettings.firstDataRow,
            SheetUtils.getLastRow(sheet),
        )
        const issues = Utils.timed(`${IssueHierarchyFormatter.name}: getting issues`, () =>
            issuesRange.getValues()
                .map(it => it[0]?.toString())
                .filter(it => it?.length)
                .filter(Utils.distinct()),
        )
        if (!issues.length) {
            return
        }

        if (GSheetProjectSettings.reorderHierarchyAutomatically) {
            Utils.timed(`${IssueHierarchyFormatter.name}: ${this.reorderIssuesAccordingToHierarchy.name}`, () =>
                this.reorderIssuesAccordingToHierarchy(issues),
            )
        }

        Utils.timed(`${IssueHierarchyFormatter.name}: ${this.formatHierarchyIssues.name}`, () =>
            this.formatHierarchyIssues(issues),
        )
    }

    static reorderAllIssuesAccordingToHierarchy() {
        this.reorderIssuesAccordingToHierarchy(undefined)
    }


    static reorderIssuesAccordingToHierarchy(issuesToReorder: string[] | undefined) {
        if (issuesToReorder != null && !issuesToReorder.length) {
            return
        }

        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName)
        ProtectionLocks.lockAllRows(sheet)

        const issuesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueColumnName)
        const childIssuesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName)

        const {
            issues,
            childIssues,
        } = SheetUtils.getColumnsStringValues(sheet, {
            issues: issuesColumn,
            childIssues: childIssuesColumn,
        }, GSheetProjectSettings.firstDataRow)

        const notEmptyIssues = issues.filter(it => it?.length)
        const notEmptyUniqueIssues = notEmptyIssues.filter(Utils.distinct())

        Utils.trimArrayEndBy(issues, it => !it?.length)
        SheetUtils.setLastRow(sheet, GSheetProjectSettings.firstDataRow + issues.length - 1)
        childIssues.length = issues.length

        if (notEmptyIssues.length === notEmptyUniqueIssues.length) {
            return GSheetProjectSettings.firstDataRow + issues.length
        }


        const moveIssues = (fromIndex, count, targetIndex) => {
            if (fromIndex === targetIndex || count <= 0) {
                return
            }

            const fromRow = GSheetProjectSettings.firstDataRow + fromIndex
            const targetRow = GSheetProjectSettings.firstDataRow + targetIndex
            if (count === 1) {
                console.info(`Moving row #${fromRow} to #${targetRow}`)
            } else {
                console.info(`Moving rows #${fromRow}-${fromRow + count - 1} to #${targetRow}`)
            }
            const range = sheet.getRange(fromRow, 1, count, 1)
            sheet.moveRows(range, targetRow)

            Utils.moveArrayElements(issues, fromIndex, count, targetIndex)
            Utils.moveArrayElements(childIssues, fromIndex, count, targetIndex)
        }

        const groupIndexes = (indexes: number[], targetIndex: number) => {
            while (indexes.length) {
                let index = indexes.shift()!
                if (index === targetIndex) {
                    continue
                }

                let bulkSize = 1
                while (indexes.length) {
                    const nextIndex = indexes[0]
                    if (nextIndex === index + bulkSize) {
                        ++bulkSize
                        indexes.shift()
                    } else {
                        break
                    }
                }

                moveIssues(index, bulkSize, targetIndex + 1)

                if (index < targetIndex) {
                    targetIndex += bulkSize - 1
                } else {
                    targetIndex += bulkSize - (index < targetIndex ? 1 : 0)
                }
            }
        }

        const hasGapsInIndexes = (indexes: number[]): boolean => {
            return indexes.length >= 2 && indexes[indexes.length - 1] - indexes[0] >= indexes.length
        }

        for (const issue of notEmptyUniqueIssues) {
            if (issuesToReorder != null && !issuesToReorder.includes(issue)) {
                continue
            }

            const getIndexes = () => issues
                .map((it, index) => issue === it ? index : null)
                .filter(index => index != null)
                .map(index => index!)
                .toSorted(Utils.numericAsc())

            const getIndexesWithoutChild = () => getIndexes()
                .filter(index => !childIssues[index]?.length)

            const getIndexesWithChild = () => getIndexes()
                .filter(index => childIssues[index]?.length)

            { // group issues without child
                const indexesWithoutChild = getIndexesWithoutChild()
                if (hasGapsInIndexes(indexesWithoutChild)) {
                    const firstIndexWithoutChild = indexesWithoutChild.shift()!
                    groupIndexes(indexesWithoutChild, firstIndexWithoutChild)
                }
            }

            { // group indexes with child
                const indexesWithChild = getIndexesWithChild()
                if (indexesWithChild.length) {
                    const indexesWithoutChild = getIndexesWithoutChild()
                    if (indexesWithoutChild.length) {
                        let targetIndex = getIndexesWithoutChild().pop()!
                        if (indexesWithChild[0] >= targetIndex) {
                            ++targetIndex
                        }
                        groupIndexes(indexesWithChild, targetIndex)

                    } else if (hasGapsInIndexes(indexesWithChild)) {
                        const firstIndexWithChild = indexesWithChild.shift()!
                        groupIndexes(indexesWithChild, firstIndexWithChild)
                    }
                }
            }
        }
    }

    static formatHierarchyIssues(issuesToFormat: string[]) {
        if (!issuesToFormat.length) {
            return
        }

        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName)
        ProtectionLocks.lockAllRows(sheet)

        const issuesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueColumnName)
        const childIssuesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName)
        const milestonesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.milestoneColumnName)
        const typesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.typeColumnName)
        const deadlinesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.deadlineColumnName)
        const titlesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.titleColumnName)
        const {
            issues,
            childIssues,
            titles,
            milestones,
            types,
            deadlines,
        } = SheetUtils.getColumnsStringValues(sheet, {
            issues: issuesColumn,
            childIssues: childIssuesColumn,
            titles: titlesColumn,
            milestones: milestonesColumn,
            types: typesColumn,
            deadlines: deadlinesColumn,
        }, GSheetProjectSettings.firstDataRow)

        const notEmptyIssues = issues.filter(it => it?.length)
        const notEmptyUniqueIssues = notEmptyIssues.filter(Utils.distinct())
        if (notEmptyIssues.length === notEmptyUniqueIssues.length) {
            return
        }

        Utils.trimArrayEndBy(issues, it => !it?.length)
        SheetUtils.setLastRow(sheet, GSheetProjectSettings.firstDataRow + issues.length - 1)
        childIssues.length = issues.length
        milestones.length = issues.length
        types.length = issues.length
        deadlines.length = issues.length


        const {
            titleFormulas,
            milestoneFormulas,
            typeFormulas,
            deadlineFormulas,
        } = SheetUtils.getColumnsFormulas(sheet, {
            titleFormulas: titlesColumn,
            milestoneFormulas: milestonesColumn,
            typeFormulas: typesColumn,
            deadlineFormulas: deadlinesColumn,
        }, GSheetProjectSettings.firstDataRow)

        titleFormulas.length = issues.length
        milestoneFormulas.length = issues.length
        typeFormulas.length = issues.length
        deadlineFormulas.length = issues.length

        const isFormulaEmptyOrDefault = (formulas: string[], index: number): boolean => {
            const formula = formulas[index]
            if (!formula?.length) {
                return true
            }

            return Utils.extractFormulaMarker(formula) === DefaultFormulas.FORMULA_MARKER
        }


        for (const issue of notEmptyUniqueIssues) {
            if (!issuesToFormat.includes(issue)) {
                continue
            }

            const getIndexes = () => issues
                .map((it, index) => issue === it ? index : null)
                .filter(index => index != null)
                .map(index => index!)
                .toSorted(Utils.numericAsc())

            const getIndexesWithoutChild = () => getIndexes()
                .filter(index => !childIssues[index]?.length)

            const getIndexesWithChild = () => getIndexes()
                .filter(index => childIssues[index]?.length)

            { // set formulas
                const indexesWithoutChild = getIndexesWithoutChild()
                const indexesWithChild = getIndexesWithChild()
                if (indexesWithoutChild.length && indexesWithChild.length) {
                    const firstIndexWithoutChild = indexesWithoutChild[0]
                    const firstRowWithoutChild = GSheetProjectSettings.firstDataRow + firstIndexWithoutChild

                    const getIssueFormula = (column: Column): string => {
                        return RangeUtils.getAbsoluteReferenceFormula(
                            sheet.getRange(firstRowWithoutChild, column),
                        )
                    }

                    const firstIndexWithChild = indexesWithChild[0]
                    const firstRowWithChild = GSheetProjectSettings.firstDataRow + firstIndexWithChild

                    sheet.getRange(firstRowWithChild, issuesColumn, indexesWithChild.length, 1)
                        .setFormula(getIssueFormula(issuesColumn))
                        .setFontSize(GSheetProjectSettings.fontSize - 2)

                    indexesWithChild.forEach(index => {
                        const row = GSheetProjectSettings.firstDataRow + index

                        if (!titles[index]?.length && isFormulaEmptyOrDefault(titleFormulas, index)) {
                            const firstTitleWithoutChildRange = sheet.getRange(firstRowWithoutChild, titlesColumn)
                            const childIssueRange = sheet.getRange(row, childIssuesColumn)
                            const formula = Utils.processFormula(`
                                =${RangeUtils.getAbsoluteA1Notation(firstTitleWithoutChildRange)}
                                & " - "
                                & ${RangeUtils.getAbsoluteA1Notation(childIssueRange)}
                            `)
                            sheet.getRange(row, titlesColumn)
                                .setFormula(formula)
                        }

                        if (!milestones[index]?.length && isFormulaEmptyOrDefault(milestoneFormulas, index)) {
                            sheet.getRange(row, milestonesColumn)
                                .setFormula(getIssueFormula(milestonesColumn))
                        }

                        if (!types[index]?.length && isFormulaEmptyOrDefault(typeFormulas, index)) {
                            sheet.getRange(row, typesColumn)
                                .setFormula(getIssueFormula(typesColumn))
                        }

                        if (!deadlines[index]?.length && isFormulaEmptyOrDefault(deadlineFormulas, index)) {
                            sheet.getRange(row, deadlinesColumn)
                                .setFormula(getIssueFormula(deadlinesColumn))
                        }
                    })
                }
            }
        }
    }

}
