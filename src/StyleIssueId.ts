class StyleIssueId {

    static formatIssueId(
        range: Range,
        issueIdsExtractor: IssueIdsExtractor,
        issueIdDecorator: IssueIdDecorator,
        issueIdToUrl: IssueIdToUrl,
        ...columnNames: string[]
    ) {
        for (const y of Utils.range(1, range.getHeight())) {
            for (const x of Utils.range(1, range.getWidth())) {
                const cell = range.getCell(y, x)
                if (!columnNames.some(name => RangeUtils.doesRangeHaveColumn(cell, name))) {
                    continue
                }

                const ids = issueIdsExtractor(cell.getValue())
                const links: Link[] = ids.map(id => {
                    return {
                        url: issueIdToUrl(id),
                        title: issueIdDecorator(id),
                    }
                })
                cell.setValue(RichTextUtils.createLinksValue(links))
            }
        }
    }

}
