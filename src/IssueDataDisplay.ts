class IssueDataDisplay extends AbstractIssueLogic {

    static reloadIssueData(range: Range) {
        const processedRange = this._processRange(range)
        if (processedRange == null) {
            return
        } else {
            range = processedRange
        }

        const sheet = range.getSheet()
        const iconColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.iconColumnName)
        const issueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName)
        const childIssueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueKeyColumnName)
        const lastDataReloadColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.lastDataReloadColumnName)
        const titleColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.titleColumnName)

        const {issues, childIssues, lastDataReload} = this._getIssueValuesWithLastReloadDate(range)
        let lastDataNotChangedCheckTimestamp: number = Date.now()
        const indexes = Array.from(Utils.range(0, issues.length - 1))
            .toSorted((i1, i2) => {
                const d1 = lastDataReload[i1]
                const d2 = lastDataReload[i2]
                if (d1 == null && d2 == null) {
                    return 0
                } else if (d1 != null && d2 != null) {
                    return d1.getTime() - d2.getTime()
                } else if (d1 != null) {
                    return 1
                } else {
                    return -1
                }
            })


        const processIndex = (index: number) => {
            const row = range.getRow() + index
            ProtectionLocks.lockRows(sheet, row)

            const cleanupColumns = (withTitle: boolean = false) => {
                const notations = [
                    [
                        withTitle ? titleColumn : null,
                        iconColumn,
                    ],
                    [
                        GSheetProjectSettings.issuesMetrics,
                        GSheetProjectSettings.counterIssuesMetrics,
                    ]
                        .flatMap(metrics => Object.keys(metrics))
                        .map(columnName => SheetUtils.findColumnByName(sheet, columnName)),
                ]
                    .flat()
                    .filter(column => column != null)
                    .map(column => sheet.getRange(row, column!))
                    .map(range => range!.getA1Notation())
                    .filter(Utils.distinct())
                if (notations.length) {
                    sheet.getRangeList(notations).setValue('')
                }

                sheet.getRange(row, lastDataReloadColumn).setValue(new Date())
            }

            if (GSheetProjectSettings.skipHiddenIssues && sheet.isRowHiddenByUser(row)) { // a slow check
                cleanupColumns()
                return
            }


            let currentIssueColumn: Column
            let originalIssueKeysText: string
            let isChildIssue = false
            if (childIssues[index]?.length) {
                currentIssueColumn = childIssueColumn
                originalIssueKeysText = childIssues[index]
                isChildIssue = true
            } else if (issues[index]?.length) {
                currentIssueColumn = issueColumn
                originalIssueKeysText = issues[index]
            } else {
                cleanupColumns(true)
                return
            }


            if (GSheetProjectSettings.notIssueKeyRegex?.test(originalIssueKeysText)) {
                cleanupColumns(true)
                return
            }


            const originalIssueKeysRange = sheet.getRange(row, currentIssueColumn)
            const isOriginalIssueKeysTextChanged = () => {
                const now = Date.now()
                const minTimestamp = now - GSheetProjectSettings.originalIssueKeysTextChangedTimeout
                if (lastDataNotChangedCheckTimestamp >= minTimestamp) {
                    return false
                }

                const currentValue = originalIssueKeysRange.getValue().toString()
                lastDataNotChangedCheckTimestamp = now

                if (currentValue !== originalIssueKeysText) {
                    Observability.reportWarning(`Content of ${originalIssueKeysRange.getA1Notation()} has been changed`)
                    return true
                }
                return false
            }


            const allIssueKeys = originalIssueKeysText
                .split(/[\r\n]+/)
                .map(key => key.trim())
                .filter(key => key.length)
                .filter(Utils.distinct())

            let issueTracker: IssueTracker | null = null
            const issueKeys = Utils.arrayOf<IssueKey>()
            const issueKeyIds = {} as Record<IssueKey, IssueId>
            const issueKeyQueries = {} as Record<IssueKey, IssueSearchQuery>
            for (let issueKey of allIssueKeys) {
                if (issueTracker != null) {
                    if (!issueTracker.supportsIssueKey(issueKey)) {
                        continue
                    }
                } else {
                    const keyTracker = GSheetProjectSettings.issueTrackers.find(it =>
                        it.supportsIssueKey(issueKey),
                    )
                    if (keyTracker != null) {
                        issueTracker = keyTracker
                    } else {
                        continue
                    }
                }

                issueKeys.push(issueKey)

                const issueId = issueTracker.extractIssueId(issueKey)
                if (issueId?.length) {
                    issueKeyIds[issueKey] = issueId
                }

                const searchQuery = issueTracker.extractSearchQuery(issueKey)
                if (searchQuery?.length) {
                    issueKeyQueries[issueKey] = searchQuery
                }
            }


            if (issueTracker == null) {
                cleanupColumns()
                return
            }


            const allIssueLinks: Link[] = allIssueKeys.map(issueKey => {
                if (issueKeys.includes(issueKey)) {
                    const issueId = issueKeyIds[issueKey]
                    if (issueId?.length) {
                        return {
                            title: issueTracker.canonizeIssueKey(issueKey),
                            url: issueTracker.getUrlForIssueId(issueId),
                        }

                    } else {
                        const searchQuery = issueKeyQueries[issueKey]
                        if (searchQuery?.length) {
                            return {
                                title: issueTracker.canonizeIssueKey(issueKey),
                                url: issueTracker.getUrlForSearchQuery(searchQuery),
                            }
                        }
                    }
                }

                return {
                    title: issueKey,
                }
            })

            if (isOriginalIssueKeysTextChanged()) {
                return
            }

            const issuesRichTextValue = RichTextUtils.createLinksValue(allIssueLinks)
            originalIssueKeysText = issuesRichTextValue.getText()
            sheet.getRange(row, currentIssueColumn).setRichTextValue(issuesRichTextValue)


            const loadedIssues: Issue[] = LazyProxy.create(() => Observability.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading issues`,
            ].join(': '), () => {
                const issueIds = Object.values(issueKeyIds).filter(Utils.distinct())
                return issueTracker?.loadIssuesByIssueId(issueIds)
            }))

            const loadedChildIssues: Issue[] = LazyProxy.create(() => Observability.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading child issues`,
            ].join(': '), () => {
                const issueIds = loadedIssues.map(it => it.id)
                return [
                    issueTracker.loadChildrenFor(loadedIssues),
                    Object.values(issueKeyQueries)
                        .filter(Utils.distinct())
                        .flatMap(query => issueTracker.searchByQuery(query)),
                ]
                    .flat()
                    .filter(Utils.distinctBy(issue => issue.id))
                    .filter(issue => !issueIds.includes(issue.id))
            }))

            const loadedBlockerIssues: Issue[] = LazyProxy.create(() => Observability.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading blocker issues`,
            ].join(': '), () => {
                const issueIds = loadedIssues.map(it => it.id)
                const allIssues = loadedIssues.concat(loadedChildIssues)
                return issueTracker.loadBlockersFor(allIssues)
                    .filter(issue => !issueIds.includes(issue.id))
            }))


            const titles = issueKeys.map(issueKey => {
                const issueId = issueKeyIds[issueKey]
                if (issueId?.length) {
                    return loadedIssues.find(issue => issue.id === issueId)?.title
                }

                if (issueKeyQueries[issueKey]?.length) {
                    return Observability.timed(
                        [
                            IssueDataDisplay.name,
                            this.reloadIssueData.name,
                            `row #${row}`,
                            `loading search title for "${issueKey}" issue key`,
                        ].join(': '),
                        () => issueTracker.loadIssueKeySearchTitle(issueKey),
                    )
                }

                return undefined
            })
                .map(title => title?.trim())
                .filter(title => title?.length)
                .map(title => title!)
            if (isOriginalIssueKeysTextChanged()) {
                return
            }
            sheet.getRange(row, titleColumn).setValue(titles.join('\n'))


            for (const handler of GSheetProjectSettings.onIssuesLoadedHandlers) {
                if (isOriginalIssueKeysTextChanged()) {
                    return
                }
                handler(
                    loadedIssues,
                    sheet,
                    row,
                    isChildIssue,
                )
            }


            for (const [columnName, issuesMetric] of Object.entries(GSheetProjectSettings.issuesMetrics)) {
                const column = SheetUtils.findColumnByName(sheet, columnName)
                if (column == null) {
                    continue
                }
                let value = issuesMetric(
                    loadedIssues,
                    loadedChildIssues,
                    loadedBlockerIssues,
                    sheet,
                    row,
                )
                if (value == null) {
                    value = ''
                } else if (Utils.isBoolean(value)) {
                    value = value ? "Yes" : ""
                }
                if (isOriginalIssueKeysTextChanged()) {
                    return
                }
                sheet.getRange(row, column).setValue(value)
            }


            for (const [columnName, issuesMetric] of Object.entries(GSheetProjectSettings.counterIssuesMetrics)) {
                const column = SheetUtils.findColumnByName(sheet, columnName)
                if (column == null) {
                    continue
                }
                const foundIssues = issuesMetric(
                    loadedIssues,
                    loadedChildIssues,
                    loadedBlockerIssues,
                    sheet,
                    row,
                )
                if (!foundIssues.length) {
                    sheet.getRange(row, column).setValue('')
                    continue
                }

                const foundIssueIds = foundIssues.map(it => it.id)
                    .filter(Utils.distinct())
                const link = {
                    title: foundIssues.length.toString(),
                    url: issueTracker.getUrlForIssueIds(foundIssueIds),
                }
                if (isOriginalIssueKeysTextChanged()) {
                    return
                }
                sheet.getRange(row, column).setRichTextValue(RichTextUtils.createLinkValue(link))
            }


            if (isOriginalIssueKeysTextChanged()) {
                return
            }
            sheet.getRange(row, lastDataReloadColumn).setValue(allIssueKeys.length ? new Date() : '')
        }


        const start = Date.now()
        let processedIndexes = 0
        for (const index of indexes) {
            const row = range.getRow() + index
            console.info(`Processing index ${index} (${++processedIndexes} / ${indexes.length}), row #${row}`)

            if (Date.now() - start >= GSheetProjectSettings.issuesLoadTimeoutMillis) {
                Observability.reportWarning("Issues load timeout occurred")
                break
            }

            const iconRange = sheet.getRange(row, iconColumn)

            try {
                Observability.timed(`loading issue data for row #${row}`, () => {
                    if (GSheetProjectSettings.loadingText?.length) {
                        iconRange.setValue(GSheetProjectSettings.loadingText)
                    } else {
                        iconRange.setFormula(`=IMAGE("${Images.loadingImageUrl}")`)
                    }
                    SpreadsheetApp.flush()

                    processIndex(index)
                })

            } catch (e) {
                Observability.reportError(`Error loading issue data for row #${row}: ${e}`, e)

            } finally {
                iconRange.setValue('')
                SpreadsheetApp.flush()
            }
        }
    }

    static reloadAllIssuesData() {
        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName)
        const range = sheet.getRange(1, 1, SheetUtils.getLastRow(sheet), SheetUtils.getLastColumn(sheet))
        this.reloadIssueData(range)
    }

}
