type Issue = Record<string, any>


interface Link {
    url: string
    title?: string
}

interface UrlWithTextOffset {
    url: string
    start: number
    end: number
}

interface SettingsRange {
    row: number
    column: number
    rows: number
    columns: number
}

type StringKeys<T> = { [k in keyof T]: T[k] extends string ? k : never }[keyof T]


type Range = GoogleAppsScript.Spreadsheet.Range
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type Protection = GoogleAppsScript.Spreadsheet.Protection
type RichTextValue = GoogleAppsScript.Spreadsheet.RichTextValue

type SheetsOnOpen = GoogleAppsScript.Events.SheetsOnOpen
type SheetsOnChange = GoogleAppsScript.Events.SheetsOnChange
type SheetsOnEdit = GoogleAppsScript.Events.SheetsOnEdit
type SheetsOnFormSubmit = GoogleAppsScript.Events.SheetsOnFormSubmit
