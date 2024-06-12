type SheetName = string
type ColumnName = string
type RangeName = string

type Row = number
type Column = number
type Formula = string

type IssuesMetric<T> = (issues: Issue[], childIssues: Issue[], blockerIssues: Issue[]) => T
type IssuesCounterMetric = (issues: Issue[], childIssues: Issue[], blockerIssues: Issue[]) => Issue[]

type Range = GoogleAppsScript.Spreadsheet.Range
type RangeList = GoogleAppsScript.Spreadsheet.RangeList
type Sheet = GoogleAppsScript.Spreadsheet.Sheet
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

type HorizontalAlignment = 'left' | 'center' | 'normal' | 'right'
type VerticalAlignment = 'top' | 'middle' | 'bottom'
type FontSize = number
