type IssueIdsExtractor = (text: string) => string[] | null
type IssueIdDecorator = (issueId: string) => string
type IssueIdToUrl = (issueId: string) => string
type IssueIdsToUrl = (issueIds: string[]) => string

type Issue = Record<string, any>
type IssuesLoader = (issueIds: string[]) => Issue[]
type IssueStringFieldGetter = (issue: Issue) => string
type IssueBooleanFieldGetter = (issue: Issue) => boolean

type IssueAggregateBooleanFieldGetter = (rootIssues: Issue[], childIssues: Issue[]) => boolean


interface Link {
    url: string
    title?: string
}

interface UrlWithTextOffset {
    url: string
    start: number
    end: number
}

interface IssueMetric {
    columnName: string
    filter: IssueBooleanFieldGetter
}

interface SettingsRange {
    row: number
    column: number
    rows: number
    columns: number
}


type Range = GoogleAppsScript.Spreadsheet.Range
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type RichTextValue = GoogleAppsScript.Spreadsheet.RichTextValue

type SheetsOnOpen = GoogleAppsScript.Events.SheetsOnOpen
type SheetsOnChange = GoogleAppsScript.Events.SheetsOnChange
type SheetsOnEdit = GoogleAppsScript.Events.SheetsOnEdit
type SheetsOnFormSubmit = GoogleAppsScript.Events.SheetsOnFormSubmit
