interface Link {
    title?: string | null
    url?: string | null
}

class RichTextUtils {

    static createLinkValue(link: Link): RichTextValue {
        return this.createLinksValue([link])
    }

    static createLinksValue(links: Link[]): RichTextValue {
        interface UrlWithTextOffset {
            url: string
            start: number
            end: number
        }

        let text = ''
        const linksWithOffsets: UrlWithTextOffset[] = []
        links.forEach(link => {
            const title = link.title?.length
                ? link.title
                : link.url
            if (!title?.length) {
                return
            }

            if (text.length) {
                text += '\n'
            }

            if (link.url?.length) {
                linksWithOffsets.push({
                    url: link.url,
                    start: text.length,
                    end: text.length + title.length,
                })
            }

            text += title
        })

        const builder = SpreadsheetApp.newRichTextValue().setText(text)
        linksWithOffsets.forEach(link => builder.setLinkUrl(link.start, link.end, link.url))
        builder.setTextStyle(SpreadsheetApp.newTextStyle()
            .setUnderline(true)
            .build(),
        )
        return builder.build()
    }

}
