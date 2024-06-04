class IssueHierarchyFormatter {

    static formatHierarchy(range: Range) {
        if (!RangeUtils.doesRangeHaveSheetColumn(
            range,
            GSheetProjectSettings.projectsSheetName,
            GSheetProjectSettings.projectsChildIssueColumnName,
        )) {
            return
        }

        const issuesRange = RangeUtils.toColumnRange(range, GSheetProjectSettings.projectsIssueColumnName)
        if (issuesRange != null) {
            this.formatHierarchyForIssues(
                issuesRange.getValues()
                    .map(it => it[0]?.toString()),
            )
        }
    }

    static formatHierarchyForAllIssues() {
        this.formatHierarchyForIssues(undefined)
    }


    static formatHierarchyForIssues(issuesToFormat: string[] | undefined) {
        issuesToFormat = issuesToFormat
            ?.filter(it => it?.length)
            ?.filter(Utils.distinct())
        if (issuesToFormat != null && !issuesToFormat.length) {
            return
        }

        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.projectsSheetName)
        ProtectionLocks.lockRowsWithProtection(sheet)

        const projectsIssueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsIssueColumnName)
        const issuesRange = SheetUtils.getColumnRange(
            sheet,
            projectsIssueColumn,
            GSheetProjectSettings.firstDataRow,
        )
        const projectsChildIssueColumn = SheetUtils.getColumnByName(
            sheet,
            GSheetProjectSettings.projectsChildIssueColumnName,
        )
        const childIssuesRange = SheetUtils.getColumnRange(
            sheet,
            projectsChildIssueColumn,
            GSheetProjectSettings.firstDataRow,
        )

        const issues = issuesRange.getValues()
            .map(it => it[0]?.toString())
            .map(it => it?.length ? it as string : null)
        Utils.trimArrayEndBy(issues, it => it == null)

        const nonNullIssues = issues.filter(it => it != null).map(it => it!)
        const nonNullUniqueIssues = nonNullIssues.filter(Utils.distinct())
        if (nonNullIssues.length === nonNullUniqueIssues.length) {
            return
        }

        const childIssues = childIssuesRange.getValues()
            .map(it => it[0]?.toString())
            .map(it => it?.length ? it as string : null)
        childIssues.length = issues.length


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

        for (const issue of nonNullUniqueIssues) {
            if (issuesToFormat != null && !issuesToFormat.includes(issue)) {
                continue
            }

            const getIndexes = () => issues
                .map((it, index) => issue === it ? index : null)
                .filter(index => index != null)
                .map(index => index!)
                .toSorted((i1, i2) => i1 - i2)

            const getIndexesWithoutChild = () => getIndexes()
                .filter(index => childIssues[index] == null)

            const getIndexesWithChild = () => getIndexes()
                .filter(index => childIssues[index] != null)

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

            { // set formulas
                const indexesWithoutChild = getIndexesWithoutChild()
                const indexesWithChild = getIndexesWithChild()
                if (indexesWithoutChild.length && indexesWithChild.length) {
                    const firstIndexWithoutChild = indexesWithoutChild[0]
                    const firstRowWithoutChild = GSheetProjectSettings.firstDataRow + firstIndexWithoutChild
                    const formula = `=${sheet.getRange(firstRowWithoutChild, projectsIssueColumn).getA1Notation()}`
                        .replace(/[A-Z]+/, '$$$&')
                        .replace(/\d+/, '$$$&')

                    const firstIndexWithChild = indexesWithChild[0]
                    const firstRowWithChild = GSheetProjectSettings.firstDataRow + firstIndexWithChild
                    const withChildRange = sheet.getRange(
                        firstRowWithChild,
                        projectsIssueColumn,
                        indexesWithChild.length,
                        1,
                    )
                    withChildRange.setFormula(formula)
                }
            }
        }
    }

}
