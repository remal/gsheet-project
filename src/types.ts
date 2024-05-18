type IssueIdsExtractor = (string: string) => string[]
type IssueIdDecorator = (string: string) => string
type IssueIdToUrl = (string: string) => string

interface Link {
    url: string
    title?: string
}

interface LinkWithOffset {
    url: string
    title?: string
    start: number
    end: number
}


type Range = GoogleAppsScript.Spreadsheet.Range
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type RichTextValue = GoogleAppsScript.Spreadsheet.RichTextValue

type SheetsOnOpen = GoogleAppsScript.Events.SheetsOnOpen
type SheetsOnChange = GoogleAppsScript.Events.SheetsOnChange
type SheetsOnEdit = GoogleAppsScript.Events.SheetsOnEdit
