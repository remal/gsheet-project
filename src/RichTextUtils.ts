class RichTextUtils {

    static createLinksValue(links: Link[]): RichTextValue {
        let text = ''
        const linksWithOffsets: LinkWithOffset[] = []
        links.forEach(link => {
            if (text.length) {
                text += '\n'
            }

            if (!link.title?.length) {
                link.title = link.url
            }

            linksWithOffsets.push({
                url: link.url,
                title: link.title,
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
