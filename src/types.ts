type IssueIdsExtractor = (text: string) => string[]
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

interface LinkWithOffset {
    url: string
    title?: string
    start: number
    end: number
}

interface IssueMetric {
    columnName: string
    filter: IssueBooleanFieldGetter
    color?: string
}


type Range = GoogleAppsScript.Spreadsheet.Range
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type RichTextValue = GoogleAppsScript.Spreadsheet.RichTextValue

type SheetsOnOpen = GoogleAppsScript.Events.SheetsOnOpen
type SheetsOnChange = GoogleAppsScript.Events.SheetsOnChange
type SheetsOnEdit = GoogleAppsScript.Events.SheetsOnEdit
type SheetsOnFormSubmit = GoogleAppsScript.Events.SheetsOnFormSubmit
