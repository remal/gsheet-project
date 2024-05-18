class IssueIdFormatter {

    private settings: GSheetProjectSettings

    constructor(settings: GSheetProjectSettings) {
        this.settings = settings;
    }


    formatIssueId(range: Range) {
        const columnNames = [
            this.settings.issueIdColumnName,
            this.settings.parentIssueIdColumnName,
        ]
        for (const y of Utils.range(1, range.getHeight())) {
            for (const x of Utils.range(1, range.getWidth())) {
                const cell = range.getCell(y, x)
                if (!columnNames.some(name => RangeUtils.doesRangeHaveColumn(cell, name))) {
                    continue
                }

                const ids = this.settings.issueIdsExtractor(cell.getValue())
                const links: Link[] = ids.map(id => {
                    return {
                        url: this.settings.issueIdToUrl(id),
                        title: this.settings.issueIdDecorator(id),
                    }
                })
                cell.setValue(RichTextUtils.createLinksValue(links))
            }
        }
    }

}
