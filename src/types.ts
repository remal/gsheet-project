type IssueId = string
type Issue = Record<string, any>

type Range = GoogleAppsScript.Spreadsheet.Range
type RangeList = GoogleAppsScript.Spreadsheet.RangeList
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type Protection = GoogleAppsScript.Spreadsheet.Protection
type NamedRange = GoogleAppsScript.Spreadsheet.NamedRange
type DataValidation = GoogleAppsScript.Spreadsheet.DataValidation
type RichTextValue = GoogleAppsScript.Spreadsheet.RichTextValue
type DeveloperMetadata = GoogleAppsScript.Spreadsheet.DeveloperMetadata

type SheetsOnOpen = GoogleAppsScript.Events.SheetsOnOpen
type SheetsOnChange = GoogleAppsScript.Events.SheetsOnChange
type SheetsOnEdit = GoogleAppsScript.Events.SheetsOnEdit
type SheetsOnFormSubmit = GoogleAppsScript.Events.SheetsOnFormSubmit

type HorizontalAlignment = 'left' | 'center' | 'normal' | 'right'
type VerticalAlignment = 'top' | 'middle' | 'bottom'

type StringKeys<T> = { [k in keyof T]: T[k] extends string ? k : never }[keyof T]
