interface Link {
    url: string
    title?: string
}

class RichTextUtils {

    static createLinksValue(links: Link[]): RichTextValue {
        interface UrlWithTextOffset {
            url: string
            start: number
            end: number
        }

        let text = ''
        const linksWithOffsets: UrlWithTextOffset[] = []
        links.forEach(link => {
            if (text.length) {
                text += '\n'
            }

            if (!link.title?.length) {
                link.title = link.url
            }

            linksWithOffsets.push({
                url: link.url,
                start: text.length,
                end: text.length + link.title.length,
            })

            text += link.title
        })

        const builder = SpreadsheetApp.newRichTextValue().setText(text)
        linksWithOffsets.forEach(link => builder.setLinkUrl(link.start, link.end, link.url))
        return builder.build()
    }

}
