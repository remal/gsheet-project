class IssueIdFormatter {

    static formatIssueId(range: Range) {
        const columnNames = [
            GSheetProjectSettings.issueIdColumnName,
            GSheetProjectSettings.parentIssueIdColumnName,
        ]
        for (const y of Utils.range(1, range.getHeight())) {
            for (const x of Utils.range(1, range.getWidth())) {
                const cell = range.getCell(y, x)
                if (!columnNames.some(name => RangeUtils.doesRangeHaveColumn(cell, name))) {
                    continue
                }

                const ids = GSheetProjectSettings.issueIdsExtractor(cell.getValue()) ?? []
                const links: Link[] = ids.map(id => {
                    return {
                        url: GSheetProjectSettings.issueIdToUrl(id),
                        title: GSheetProjectSettings.issueIdDecorator(id),
                    }
                })
                cell.setRichTextValue(RichTextUtils.createLinksValue(links))
            }
        }
    }

}
