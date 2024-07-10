type SheetName = string
type ColumnName = string
type RangeName = string

type Row = number
type Column = number
type Formula = string

type OnIssuesLoaded = (
    issues: Issue[],
    sheet: Sheet,
    row: Row,
) => void

type IssuesMetric<T> = (
    issues: Issue[],
    childIssues: Issue[],
    blockerIssues: Issue[],
    sheet: Sheet,
    row: Row,
) => T
type IssuesCounterMetric = (
    issues: Issue[],
    childIssues: Issue[],
    blockerIssues: Issue[],
    sheet: Sheet,
    row: Row,
) => Issue[]

type Range = GoogleAppsScript.Spreadsheet.Range
type RangeList = GoogleAppsScript.Spreadsheet.RangeList
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
type Filter = GoogleAppsScript.Spreadsheet.Filter
type FilterCriteria = GoogleAppsScript.Spreadsheet.FilterCriteria
type Protection = GoogleAppsScript.Spreadsheet.Protection
type NamedRange = GoogleAppsScript.Spreadsheet.NamedRange
type DataValidation = GoogleAppsScript.Spreadsheet.DataValidation
type RichTextValue = GoogleAppsScript.Spreadsheet.RichTextValue
type DeveloperMetadata = GoogleAppsScript.Spreadsheet.DeveloperMetadata
type ConditionalFormatRule = GoogleAppsScript.Spreadsheet.ConditionalFormatRule
type ConditionalFormatRuleBuilder = GoogleAppsScript.Spreadsheet.ConditionalFormatRuleBuilder

type SheetsOnOpen = GoogleAppsScript.Events.SheetsOnOpen
type SheetsOnChange = GoogleAppsScript.Events.SheetsOnChange
type SheetsOnEdit = GoogleAppsScript.Events.SheetsOnEdit
type SheetsOnFormSubmit = GoogleAppsScript.Events.SheetsOnFormSubmit

type Lock = GoogleAppsScript.Lock.Lock

type HorizontalAlignment = 'left' | 'center' | 'normal' | 'right'
type VerticalAlignment = 'top' | 'middle' | 'bottom'
type FontSize = number
type Color = string
