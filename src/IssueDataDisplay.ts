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
        const issueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueColumnName)
        const childIssueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName)
        const lastDataReloadColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.lastDataReloadColumnName)
        const titleColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.titleColumnName)

        const {issues, childIssues, lastDataReload} = this._getIssueValuesWithLastReloadDate(range)
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

        const start = Date.now()
        for (const index of indexes) {
            if (Date.now() - start >= GSheetProjectSettings.issuesLoadTimeoutMillis) {
                const message = "Issues load timeout occurred"
                console.warn(message)
                //SpreadsheetApp.getActiveSpreadsheet().toast(message)
                break
            }

            const row = range.getRow() + index

            const cleanupColumns = () => {
                const notations = [
                    [
                        sheet.getRange(row, titleColumn),
                        sheet.getRange(row, iconColumn),
                    ],
                    Object.keys(GSheetProjectSettings.booleanIssuesMetrics)
                        .map(columnName => SheetUtils.findColumnByName(sheet, columnName))
                        .filter(column => column != null)
                        .map(column => sheet.getRange(row, column!)),
                    Object.keys(GSheetProjectSettings.counterIssuesMetrics)
                        .map(columnName => SheetUtils.findColumnByName(sheet, columnName))
                        .filter(column => column != null)
                        .map(column => sheet.getRange(row, column!)),
                ]
                    .flat()
                    .map(range => range.getA1Notation())
                if (notations.length) {
                    sheet.getRangeList(notations).setValue('')
                }

                sheet.getRange(row, lastDataReloadColumn).setValue(new Date())
            }

            if (GSheetProjectSettings.skipHiddenIssues && sheet.isRowHiddenByUser(row)) { // a slow check
                cleanupColumns()
                continue
            }

            if (GSheetProjectSettings.useLoadingImage) {
                sheet.getRange(row, iconColumn).setFormula(`=IMAGE("${Images.loadingImageUrl}")`)
            } else {
                sheet.getRange(row, iconColumn).setValue('...')
            }


            let currentIssueColumn: Column
            let originalIssueKeysText: string
            if (childIssues[index]?.length) {
                currentIssueColumn = childIssueColumn
                originalIssueKeysText = childIssues[index]
            } else if (issues[index]?.length) {
                currentIssueColumn = issueColumn
                originalIssueKeysText = issues[index]
            } else {
                cleanupColumns()
                continue
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
                continue
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

            sheet.getRange(row, currentIssueColumn).setRichTextValue(RichTextUtils.createLinksValue(allIssueLinks))


            const loadedIssues: Issue[] = LazyProxy.create(() => Utils.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading issues`,
            ].join(': '), () => {
                const issueIds = Object.values(issueKeyIds).filter(Utils.distinct())
                return issueTracker.loadIssues(issueIds)
            }))

            const loadedChildIssues: Issue[] = LazyProxy.create(() => Utils.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading child issues`,
            ].join(': '), () => {
                const issueIds = loadedIssues.map(it => it.id)
                return [
                    issueTracker.loadChildren(issueIds),
                    Object.values(issueKeyQueries)
                        .filter(Utils.distinct())
                        .flatMap(query => issueTracker.search(query)),
                ]
                    .flat()
                    .filter(Utils.distinctBy(issue => issue.id))
                    .filter(issue => !issueIds.includes(issue.id))
            }))

            const loadedBlockerIssues: Issue[] = LazyProxy.create(() => Utils.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading blocker issues`,
            ].join(': '), () => {
                const allIssueIds = [loadedIssues, loadedChildIssues]
                    .flatMap(it => it.map(it => it.id))
                    .filter(Utils.distinct())
                return issueTracker.loadBlockers(allIssueIds)
                    .filter(issue => !allIssueIds.includes(issue.id))
            }))


            const titles = issueKeys.map(issueKey => {
                const issueId = issueKeyIds[issueKey]
                if (issueId?.length) {
                    return loadedIssues.find(issue => issue.id)?.title
                }

                if (issueKeyQueries[issueKey]?.length) {
                    return issueTracker.loadIssueKeySearchTitle(issueKey)
                }

                return undefined
            })
                .map(title => title?.trim())
                .filter(title => title?.length)
                .map(title => title!)
            if (titles.length) {
                sheet.getRange(row, titleColumn).setValue(titles.join('\n'))
            }


            for (const [columnName, issuesMetric] of Object.entries(GSheetProjectSettings.booleanIssuesMetrics)) {
                const column = SheetUtils.findColumnByName(sheet, columnName)
                if (column == null) {
                    continue
                }
                const value = issuesMetric(loadedIssues, loadedChildIssues, loadedBlockerIssues)
                sheet.getRange(row, column).setValue(
                    value ? "Yes" : '',
                )
            }


            for (const [columnName, issuesCounterMetric] of Object.entries(GSheetProjectSettings.counterIssuesMetrics)) {
                const column = SheetUtils.findColumnByName(sheet, columnName)
                if (column == null) {
                    continue
                }
                const foundIssues = issuesCounterMetric(loadedIssues, loadedChildIssues, loadedBlockerIssues)
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
                sheet.getRange(row, column).setRichTextValue(RichTextUtils.createLinkValue(link))
            }


            sheet.getRange(row, lastDataReloadColumn).setValue(allIssueKeys.length ? new Date() : '')
            sheet.getRange(row, iconColumn).setValue('')
        }
    }

    static reloadAllIssuesData() {
        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName)
        const range = sheet.getRange(1, 1, SheetUtils.getLastRow(sheet), SheetUtils.getLastColumn(sheet))
        this.reloadIssueData(range)
    }

}
