/**
 * SHA-256 digest of the provided input
 * @param {unknown} value
 * @returns {string}
 * @customFunction
 */
function SHA256(value) {
    const string = value?.toString() ?? '';
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, string);
    return digest
        .map(num => num < 0 ? num + 256 : num)
        .map(num => num.toString(16))
        .map(num => (num.length === 1 ? '0' : '') + num)
        .join('');
}
function refreshSelectedRowsOfGSheetProject() {
    const range = SpreadsheetApp.getActiveRange();
    if (range == null) {
        return;
    }
    const sheet = range.getSheet();
    if (!SheetUtils.isGridSheet(sheet)) {
        return;
    }
    EntryPoint.entryPoint(() => {
        const rowsRange = sheet.getRange(`${range.getRow()}:${range.getRow() + range.getNumRows() - 1}`);
        onEditGSheetProject({
            range: rowsRange,
        });
    });
}
function refreshAllRowsOfGSheetProject() {
    EntryPoint.entryPoint(() => {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
            const rowsRange = sheet.getRange(`1:${SheetUtils.getLastRow(sheet)}`);
            onEditGSheetProject({
                range: rowsRange,
            });
        });
    }, false);
}
function reapplyDefaultFormulasOfGSheetProject() {
    EntryPoint.entryPoint(() => {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
            const rowsRange = sheet.getRange(`1:${SheetUtils.getLastRow(sheet)}`);
            DefaultFormulas.insertDefaultFormulas(rowsRange, true);
        });
    });
}
function applyDefaultStylesOfGSheetProject() {
    EntryPoint.entryPoint(() => {
        SheetLayouts.migrate();
    });
}
function cleanupGSheetProject() {
    EntryPoint.entryPoint(() => {
        SheetLayouts.migrateIfNeeded();
        ConditionalFormatting.removeDuplicateConditionalFormatRules();
        ConditionalFormatting.combineConditionalFormatRules();
        ProtectionLocks.releaseExpiredLocks();
        PropertyLocks.releaseExpiredPropertyLocks();
    }, false);
}
function onOpenGSheetProject(event) {
    EntryPoint.entryPoint(() => {
        SheetLayouts.migrateIfNeeded();
    }, false);
    SpreadsheetApp.getUi()
        .createMenu("GSheetProject")
        .addItem("Refresh selected rows", refreshSelectedRowsOfGSheetProject.name)
        .addItem("Refresh all rows", refreshAllRowsOfGSheetProject.name)
        .addItem("Reapply default formulas", reapplyDefaultFormulasOfGSheetProject.name)
        .addItem("Apply default styles", applyDefaultStylesOfGSheetProject.name)
        .addToUi();
}
function onChangeGSheetProject(event) {
    function onInsert() {
        EntryPoint.entryPoint(() => {
            CommonFormatter.applyCommonFormatsToAllSheets();
        });
    }
    function onRemove() {
        applyDefaultStylesOfGSheetProject();
    }
    const changeType = event?.changeType?.toString() ?? '';
    if (['INSERT_ROW', 'INSERT_COLUMN'].includes(changeType)) {
        onInsert();
    }
    else if (['REMOVE_COLUMN'].includes(changeType)) {
        onRemove();
    }
}
function onEditGSheetProject(event) {
    const range = event?.range;
    if (range == null) {
        return;
    }
    EntryPoint.entryPoint(() => {
        Observability.timed(`Common format`, () => CommonFormatter.applyCommonFormatsToRowRange(range));
        //Observability.timed(`Done logic`, () => DoneLogic.executeDoneLogic(range))
        Observability.timed(`Default formulas`, () => DefaultFormulas.insertDefaultFormulas(range));
        Observability.timed(`Issue hierarchy`, () => IssueHierarchyFormatter.formatHierarchy(range));
        Observability.timed(`Reload issue data`, () => IssueDataDisplay.reloadIssueData(range));
    });
}
function onFormSubmitGSheetProject(event) {
    onEditGSheetProject({
        range: event?.range,
    });
}
var _a;
class GSheetProjectSettings {
    static computeStringSettingsHash() {
        const hashableValues = {};
        const keys = Object.keys(_a).toSorted();
        for (const key of keys) {
            let value = _a[key];
            if (value instanceof RegExp) {
                value = value.toString();
            }
            if (value == null
                || typeof value === 'function'
                || typeof value === 'object') {
                continue;
            }
            hashableValues[key] = value;
        }
        const json = JSON.stringify(hashableValues);
        return SHA256(json);
    }
}
_a = GSheetProjectSettings;
GSheetProjectSettings.titleRow = 1;
GSheetProjectSettings.firstDataRow = _a.titleRow + 1;
GSheetProjectSettings.lockColumns = false;
GSheetProjectSettings.lockRows = false;
GSheetProjectSettings.updateConditionalFormatRules = true;
GSheetProjectSettings.reorderHierarchyAutomatically = false;
GSheetProjectSettings.skipHiddenIssues = true;
//static restoreUndoneEnd: boolean = false
GSheetProjectSettings.issuesRangeName = 'Issues';
GSheetProjectSettings.childIssuesRangeName = 'ChildIssues';
GSheetProjectSettings.milestonesRangeName = "Milestones";
GSheetProjectSettings.titlesRangeName = "Titles";
GSheetProjectSettings.teamsRangeName = "Teams";
GSheetProjectSettings.estimatesRangeName = "Estimates";
GSheetProjectSettings.startsRangeName = "Starts";
GSheetProjectSettings.endsRangeName = "Ends";
GSheetProjectSettings.earliestStartsRangeName = "EarliestStarts";
GSheetProjectSettings.deadlinesRangeName = "Deadlines";
GSheetProjectSettings.inProgressesRangeName = undefined;
GSheetProjectSettings.codeCompletesRangeName = undefined;
GSheetProjectSettings.settingsScheduleStartRangeName = 'ScheduleStart';
GSheetProjectSettings.settingsScheduleBufferRangeName = 'ScheduleBuffer';
GSheetProjectSettings.settingsTeamsTableRangeName = 'TeamsTable';
GSheetProjectSettings.settingsTeamsTableTeamRangeName = 'TeamsTableTeam';
GSheetProjectSettings.settingsTeamsTableResourcesRangeName = 'TeamsTableResources';
GSheetProjectSettings.settingsMilestonesTableRangeName = 'MilestonesTable';
GSheetProjectSettings.settingsMilestonesTableMilestoneRangeName = 'MilestonesTableMilestone';
GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName = 'MilestonesTableDeadline';
GSheetProjectSettings.publicHolidaysRangeName = 'PublicHolidays';
GSheetProjectSettings.notIssueKeyRegex = /^\s*\W/;
GSheetProjectSettings.bufferIssueKeyRegex = /^(buffer|reserve)/i;
GSheetProjectSettings.issueTrackers = [];
GSheetProjectSettings.issuesLoadTimeoutMillis = 5 * 60 * 1000;
GSheetProjectSettings.onIssuesLoadedHandlers = [];
GSheetProjectSettings.issuesMetrics = {};
GSheetProjectSettings.counterIssuesMetrics = {};
GSheetProjectSettings.originalIssueKeysTextChangedTimeout = 250;
GSheetProjectSettings.useLockService = true;
GSheetProjectSettings.lockTimeoutMillis = 5 * 60 * 1000;
GSheetProjectSettings.sheetName = "Projects";
GSheetProjectSettings.iconColumnName = "icon";
//static doneColumnName: ColumnName = "Done"
GSheetProjectSettings.milestoneColumnName = "Milestone";
GSheetProjectSettings.typeColumnName = "Type";
GSheetProjectSettings.issueKeyColumnName = "Issue";
GSheetProjectSettings.childIssueKeyColumnName = "Child\nIssue";
GSheetProjectSettings.lastDataReloadColumnName = "Last\nReload";
GSheetProjectSettings.titleColumnName = "Title";
GSheetProjectSettings.teamColumnName = "Team";
GSheetProjectSettings.estimateColumnName = "Estimate\n(days)";
GSheetProjectSettings.earliestStartColumnName = "Earliest\nStart";
GSheetProjectSettings.deadlineColumnName = "Deadline";
GSheetProjectSettings.startColumnName = "Start";
GSheetProjectSettings.endColumnName = "End";
//static issueHashColumnName: ColumnName = "Issue Hash"
GSheetProjectSettings.settingsSheetName = "Settings";
GSheetProjectSettings.loadingText = '\u2B6E'; // alternative: '\uD83D\uDD03'
GSheetProjectSettings.indent = 4;
GSheetProjectSettings.fontSize = 10;
// see https://spreadsheet.dev/how-to-get-the-hexadecimal-codes-of-colors-in-google-sheets
GSheetProjectSettings.errorColor = '#ff0000';
GSheetProjectSettings.importantWarningColor = '#e06666';
GSheetProjectSettings.warningColor = '#e69138';
GSheetProjectSettings.unimportantWarningColor = '#fce5cd';
GSheetProjectSettings.unimportantColor = '#b7b7b7';
class AbstractIssueLogic {
    static _processRange(range) {
        if (![
            GSheetProjectSettings.issueKeyColumnName,
            GSheetProjectSettings.childIssueKeyColumnName,
            GSheetProjectSettings.teamColumnName,
        ].some(columnName => RangeUtils.doesRangeHaveSheetColumn(range, GSheetProjectSettings.sheetName, columnName))) {
            return null;
        }
        const sheet = range.getSheet();
        ProtectionLocks.lockAllColumns(sheet);
        range = RangeUtils.withMinMaxRows(range, GSheetProjectSettings.firstDataRow, SheetUtils.getLastRow(sheet));
        const startRow = range.getRow();
        const rows = range.getNumRows();
        const endRow = startRow + rows - 1;
        ProtectionLocks.lockRows(sheet, endRow);
        return range;
    }
    static _getIssueValues(range) {
        const sheet = range.getSheet();
        const startRow = range.getRow();
        const endRow = startRow + range.getNumRows() - 1;
        const result = SheetUtils.getColumnsStringValues(sheet, {
            issues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName),
            childIssues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueKeyColumnName),
        }, startRow, endRow);
        Utils.trimArrayEndBy(result.issues, it => !it?.length);
        result.childIssues.length = result.issues.length;
        return result;
    }
    static _getIssueValuesWithLastReloadDate(range) {
        const sheet = range.getSheet();
        const startRow = range.getRow();
        const endRow = startRow + range.getNumRows() - 1;
        const result = SheetUtils.getColumnsValues(sheet, {
            issues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName),
            childIssues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueKeyColumnName),
            lastDataReload: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.lastDataReloadColumnName),
        }, startRow, endRow);
        Utils.trimArrayEndBy(result.issues, it => !it?.toString()?.length);
        result.childIssues.length = result.issues.length;
        result.lastDataReload.length = result.issues.length;
        return {
            issues: result.issues.map(it => it?.toString()),
            childIssues: result.childIssues.map(it => it?.toString()),
            lastDataReload: result.lastDataReload.map(it => Utils.parseDate(it)),
        };
    }
    static _getValues(range, column) {
        return RangeUtils.toColumnRange(range, column).getValues()
            .map(it => it[0]);
    }
    static _getStringValues(range, column) {
        return this._getValues(range, column).map(it => it.toString());
    }
    static _getFormulas(range, column) {
        return RangeUtils.toColumnRange(range, column).getFormulas()
            .map(it => it[0]);
    }
}
class CommonFormatter {
    static applyCommonFormatsToAllSheets() {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
            this.highlightCellsWithFormula(sheet);
            const range = SheetUtils.getWholeSheetRange(sheet);
            this.applyCommonFormatsToRange(range);
        });
    }
    static highlightCellsWithFormula(sheet) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet);
        }
        const range = SheetUtils.getWholeSheetRange(sheet);
        ConditionalFormatting.addConditionalFormatRule(range, {
            scope: 'common',
            order: 10000,
            configurer: builder => builder
                .whenFormulaSatisfied(`=
                        ISFORMULA(A1)
                    `)
                .setItalic(true),
            //.setFontColor('#333'),
        }, false);
    }
    static applyCommonFormatsToRowRange(range) {
        const sheet = range.getSheet();
        range = sheet.getRange(range.getRow(), 1, range.getNumRows(), SheetUtils.getMaxColumns(sheet));
        this.applyCommonFormatsToRange(range);
    }
    static applyCommonFormatsToRange(range) {
        range
            .setVerticalAlignment('middle');
    }
}
class ConditionalFormatRuleUtils {
    static extractFormula(rule) {
        const condition = rule.getBooleanCondition();
        if (condition == null) {
            return undefined;
        }
        if (condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
            return undefined;
        }
        return condition.getCriteriaValues()[0].toString();
    }
}
class ConditionalFormatting {
    static addConditionalFormatRule(range, orderedRule, addIsFormulaRule = true) {
        if (!GSheetProjectSettings.updateConditionalFormatRules) {
            return;
        }
        if ((orderedRule.order | 0) !== orderedRule.order) {
            throw new Error(`Order is not integer: ${orderedRule.order}`);
        }
        if (orderedRule.order <= 0) {
            throw new Error(`Order is <= 0: ${orderedRule.order}`);
        }
        const builder = SpreadsheetApp.newConditionalFormatRule();
        builder.setRanges([range]);
        orderedRule.configurer(builder);
        let formula = ConditionalFormatRuleUtils.extractFormula(builder);
        if (formula == null) {
            throw new Error(`Not a boolean condition with formula`);
        }
        formula = Formulas.processFormula(formula)
            .replace(/^=+/, '');
        const newRuleFormula = Formulas.processFormula(`=
            AND(
                ${formula},
                "GSPs"<>"${orderedRule.scope}",
                "GSPo"<>"${orderedRule.order + 0.2}"
            )
        `);
        builder.whenFormulaSatisfied(newRuleFormula);
        const newRule = builder.build();
        const newRules = [newRule];
        if (addIsFormulaRule) {
            const newIsFormula = Formulas.processFormula(`=
                    AND(
                        ISFORMULA(#SELF),
                        ${formula},
                        "GSPs"<>"${orderedRule.scope}",
                        "GSPo"<>"${orderedRule.order + 0.1}"
                    )
                `);
            const newIsFormulaRule = newRule.copy()
                .whenFormulaSatisfied(newIsFormula)
                .setItalic(true)
                .build();
            newRules.push(newIsFormulaRule);
        }
        const sheet = range.getSheet();
        let rules = sheet.getConditionalFormatRules() ?? [];
        rules = rules.filter(rule => !(this._extractScope(rule) === orderedRule.scope && this._extractIntOrder(rule) === orderedRule.order));
        rules.push(...newRules);
        rules = rules.toSorted((r1, r2) => {
            const o1 = this._extractFloatOrder(r1) ?? 0;
            const o2 = this._extractFloatOrder(r2) ?? 0;
            return o1 - o2;
        });
        sheet.setConditionalFormatRules(rules);
    }
    static removeConditionalFormatRulesByScope(sheet, scopeToRemove) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet);
        }
        const rules = sheet.getConditionalFormatRules() ?? [];
        const filteredRules = rules.filter(rule => this._extractScope(rule) !== scopeToRemove);
        if (filteredRules.length !== rules.length) {
            sheet.setConditionalFormatRules(filteredRules);
        }
    }
    static removeDuplicateConditionalFormatRules(sheet) {
        if (sheet == null) {
            SpreadsheetApp.getActiveSpreadsheet().getSheets()
                .filter(sheet => SheetUtils.isGridSheet(sheet))
                .forEach(sheet => this.removeDuplicateConditionalFormatRules(sheet));
            return;
        }
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet);
        }
        const rules = sheet.getConditionalFormatRules() ?? [];
        const filteredRules = rules.filter(Utils.distinctBy(rule => JSON.stringify(Utils.toJsonObject(rule))));
        if (filteredRules.length !== rules.length) {
            sheet.setConditionalFormatRules(filteredRules);
        }
    }
    static combineConditionalFormatRules(sheet) {
        if (sheet == null) {
            SpreadsheetApp.getActiveSpreadsheet().getSheets()
                .filter(sheet => SheetUtils.isGridSheet(sheet))
                .forEach(sheet => this.combineConditionalFormatRules(sheet));
            return;
        }
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet);
        }
        const rules = sheet.getConditionalFormatRules() ?? [];
        if (rules.length <= 1) {
            return;
        }
        const originalRules = [...rules];
        const isMergeableRule = (rule) => {
            const ranges = rule.getRanges();
            const firstRange = ranges.shift();
            if (firstRange == null) {
                return false;
            }
            for (const range of ranges) {
                if (range.getColumn() !== firstRange.getColumn()
                    || range.getNumColumns() !== firstRange.getNumColumns()) {
                    return false;
                }
            }
            return true;
        };
        const getRuleKey = (rule) => {
            const jsonObject = Utils.toJsonObject(rule);
            delete jsonObject['ranges'];
            const ranges = rule.getRanges();
            const firstRange = ranges.shift();
            jsonObject['columns'] = Array.from(Utils.range(firstRange.getColumn(), firstRange.getColumn() + firstRange.getNumColumns() - 1));
            return JSON.stringify(jsonObject);
        };
        for (let index = 0; index < rules.length - 1; ++index) {
            const rule = rules[index];
            if (!isMergeableRule(rule)) {
                continue;
            }
            const ruleKey = getRuleKey(rule);
            let similarRule = null;
            for (let otherIndex = index + 1; otherIndex < rules.length; ++otherIndex) {
                const otherRule = rules[otherIndex];
                if (!isMergeableRule(otherRule)) {
                    continue;
                }
                const otherRuleKey = getRuleKey(otherRule);
                if (otherRuleKey === ruleKey) {
                    similarRule = otherRule;
                    rules.splice(otherIndex, 1);
                    break;
                }
            }
            if (similarRule == null) {
                continue;
            }
            const ranges = [...rule.getRanges(), ...similarRule.getRanges()];
            let newRanges = [...ranges];
            newRanges = newRanges.toSorted((r1, r2) => {
                const row1 = r1.getRow();
                const row2 = r2.getRow();
                if (row1 === row2) {
                    return r2.getNumRows() - r1.getNumRows();
                }
                return row1 - row2;
            });
            for (let rangeIndex = 0; rangeIndex < newRanges.length - 1; ++rangeIndex) {
                let range = newRanges[rangeIndex];
                let firstRow = range.getRow();
                let lastRow = firstRow + range.getNumRows() - 1;
                for (let nextRangeIndex = rangeIndex + 1; nextRangeIndex < newRanges.length; ++nextRangeIndex) {
                    const nextRange = newRanges[nextRangeIndex];
                    const nextFirstRow = nextRange.getRow();
                    if (nextFirstRow <= lastRow) {
                        const nextLastRow = nextFirstRow + nextRange.getNumRows() - 1;
                        firstRow = Math.min(firstRow, nextFirstRow);
                        lastRow = Math.max(lastRow, nextLastRow);
                        lastRow = Math.min(lastRow, SheetUtils.getMaxRows(sheet));
                        newRanges[rangeIndex] = range = range.getSheet().getRange(firstRow, range.getColumn(), lastRow - firstRow + 1, range.getNumColumns());
                        newRanges.splice(nextRangeIndex, 1);
                        --nextRangeIndex;
                    }
                }
            }
            console.warn([
                ConditionalFormatting.name,
                `Combining ${ranges.map(it => it.getA1Notation())} into ${newRanges.map(it => it.getA1Notation())}`,
                ruleKey,
            ].join(': '));
            rules[index] = rule.copy().setRanges(newRanges).build();
        }
        if (rules.length !== originalRules.length) {
            sheet.setConditionalFormatRules(rules);
        }
    }
    static _extractScope(rule) {
        if (!Utils.isString(rule)) {
            const formula = ConditionalFormatRuleUtils.extractFormula(rule);
            if (formula == null) {
                return undefined;
            }
            rule = formula;
        }
        const match = rule.match(/"GSPs"\s*<>\s*"([^"]*)"/);
        if (match) {
            return match[1];
        }
        return undefined;
    }
    static _extractIntOrder(rule) {
        if (!Utils.isString(rule)) {
            const formula = ConditionalFormatRuleUtils.extractFormula(rule);
            if (formula == null) {
                return undefined;
            }
            rule = formula;
        }
        const match = rule.match(/"GSPo"\s*<>\s*"(\d+)(\.\d*)?"/);
        if (match) {
            return parseInt(match[1]);
        }
        return undefined;
    }
    static _extractFloatOrder(rule) {
        if (!Utils.isString(rule)) {
            const formula = ConditionalFormatRuleUtils.extractFormula(rule);
            if (formula == null) {
                return undefined;
            }
            rule = formula;
        }
        const match = rule.match(/"GSPo"\s*<>\s*"(\d+(\.\d*)?)"/);
        if (match) {
            return parseFloat(match[1]);
        }
        return undefined;
    }
    static _ruleKey(rule) {
        const jsonObject = Utils.toJsonObject(rule);
        delete jsonObject['ranges'];
        return JSON.stringify(jsonObject);
    }
}
class DefaultFormulas extends AbstractIssueLogic {
    static isDefaultFormula(formula) {
        return Formulas.extractFormulaMarkers(formula).includes(this._DEFAULT_FORMULA_MARKER);
    }
    static isDefaultChildFormula(formula) {
        return Formulas.extractFormulaMarkers(formula).includes(this._DEFAULT_CHILD_FORMULA_MARKER);
    }
    static insertDefaultFormulas(range, rewriteExistingDefaultFormula = false) {
        const processedRange = this._processRange(range);
        if (processedRange == null) {
            return;
        }
        else {
            range = processedRange;
        }
        const sheet = range.getSheet();
        const startRow = range.getRow();
        const rows = range.getNumRows();
        const endRow = startRow + rows - 1;
        const { issues, childIssues } = this._getIssueValues(sheet.getRange(GSheetProjectSettings.firstDataRow, range.getColumn(), endRow - GSheetProjectSettings.firstDataRow + 1, range.getNumColumns()));
        const getParentIssueRow = (issueIndex) => {
            const issue = issues[issueIndex];
            if (!issue?.length) {
                return undefined;
            }
            const index = issues.indexOf(issue);
            if (index < 0) {
                return undefined;
            }
            const childIssue = childIssues[index];
            if (childIssue?.length) {
                return undefined;
            }
            return GSheetProjectSettings.firstDataRow + index;
        };
        const issueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName);
        const childIssueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueKeyColumnName);
        const childIssueFormulas = LazyProxy.create(() => SheetUtils.getColumnsFormulas(sheet, { childIssues: childIssueColumn }, startRow, endRow).childIssues);
        const milestoneColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.milestoneColumnName);
        const typeColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.typeColumnName);
        const titleColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.titleColumnName);
        const teamColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.teamColumnName);
        const estimateColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName);
        const startColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.startColumnName);
        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName);
        const earliestStartColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.earliestStartColumnName);
        const deadlineColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.deadlineColumnName);
        const allValuesColumns = {};
        [
            milestoneColumn,
            typeColumn,
            titleColumn,
            teamColumn,
            estimateColumn,
            startColumn,
            endColumn,
            earliestStartColumn,
            deadlineColumn,
        ].forEach(column => allValuesColumns[column.toString()] = column);
        const allValues = LazyProxy.create(() => SheetUtils.getColumnsStringValues(sheet, allValuesColumns, startRow, endRow));
        const getValues = (column) => {
            if (column === issueColumn) {
                return issues.slice(-rows);
            }
            else if (column === childIssueColumn) {
                return childIssues.slice(-rows);
            }
            return allValues[column.toString()] ?? (() => {
                throw new Error(`Column ${column} is not prefetched`);
            })();
        };
        const allFormulas = LazyProxy.create(() => SheetUtils.getColumnsFormulas(sheet, allValuesColumns, startRow, endRow));
        const getFormulas = (column) => {
            if (column === childIssueColumn) {
                return childIssueFormulas;
            }
            return allFormulas[column.toString()] ?? (() => {
                throw new Error(`Column ${column} is not prefetched`);
            })();
        };
        const addFormulas = (column, formulaGenerator) => {
            const values = getValues(column);
            const formulas = getFormulas(column);
            for (let row = startRow; row <= endRow; ++row) {
                const issueIndex = row - GSheetProjectSettings.firstDataRow;
                const issue = issues[issueIndex];
                const childIssue = childIssues[issueIndex];
                const index = row - startRow;
                let value = values[index];
                let formula = formulas[index];
                if (GSheetProjectSettings.notIssueKeyRegex?.test(issue ?? '')
                    || (!issue?.length && !childIssue?.length)) {
                    if (formula?.length) {
                        sheet.getRange(row, column).setFormula('');
                    }
                    continue;
                }
                const isChild = !!childIssue?.length;
                const isDefaultFormula = this.isDefaultFormula(formula);
                const isDefaultChildFormula = this.isDefaultChildFormula(formula);
                if ((isChild && isDefaultFormula)
                    || (!isChild && isDefaultChildFormula)
                    || (rewriteExistingDefaultFormula && (isDefaultFormula || isDefaultChildFormula))) {
                    value = '';
                    formula = '';
                }
                if (!value?.length && !formula?.length) {
                    console.info([
                        DefaultFormulas.name,
                        sheet.getSheetName(),
                        addFormulas.name,
                        `column #${column}`,
                        `row #${row}`,
                    ].join(': '));
                    const isBuffer = !!GSheetProjectSettings.bufferIssueKeyRegex?.test(issue ?? '');
                    let formula = Formulas.processFormula(formulaGenerator(row, isBuffer, isChild, issueIndex, index) ?? '');
                    if (formula.length) {
                        formula = Formulas.addFormulaMarkers(formula, isChild ? this._DEFAULT_CHILD_FORMULA_MARKER : this._DEFAULT_FORMULA_MARKER, isBuffer ? this._DEFAULT_BUFFER_FORMULA_MARKER : null);
                        sheet.getRange(row, column).setFormula(formula);
                    }
                    else {
                        sheet.getRange(row, column).setFormula('');
                    }
                }
            }
        };
        addFormulas(childIssueColumn, (row, isBuffer, isChild, issueIndex) => {
            if (isBuffer) {
                childIssues[issueIndex] = `placeholder: ${addFormulas.name}`;
                return `=
                    IF(
                        #SELF_COLUMN(${GSheetProjectSettings.teamsRangeName}) <> "",
                        #SELF_COLUMN(${GSheetProjectSettings.teamsRangeName})
                        & " - "
                        & COUNTIFS(
                            OFFSET(
                                ${GSheetProjectSettings.issuesRangeName},
                                0,
                                0,
                                ROW() - ${GSheetProjectSettings.firstDataRow} + 1,
                                1
                            ), "="&#SELF_COLUMN(${GSheetProjectSettings.issuesRangeName}),
                            OFFSET(
                                ${GSheetProjectSettings.teamsRangeName},
                                0,
                                0,
                                ROW() - ${GSheetProjectSettings.firstDataRow} + 1,
                                1
                            ), "="&#SELF_COLUMN(${GSheetProjectSettings.teamsRangeName})
                        ),
                        ""
                    )
                `;
            }
            return undefined;
        });
        addFormulas(milestoneColumn, (row, isBuffer, isChild, issueIndex) => {
            if (isChild) {
                const parentIssueRow = getParentIssueRow(issueIndex);
                if (parentIssueRow != null) {
                    const milestoneA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(parentIssueRow, milestoneColumn));
                    return `=${milestoneA1Notation}`;
                }
            }
            return undefined;
        });
        addFormulas(typeColumn, (row, isBuffer, isChild, issueIndex) => {
            if (isChild) {
                const parentIssueRow = getParentIssueRow(issueIndex);
                if (parentIssueRow != null) {
                    const typeA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(parentIssueRow, typeColumn));
                    return `=${typeA1Notation}`;
                }
            }
            return undefined;
        });
        addFormulas(titleColumn, (row, isBuffer, isChild, issueIndex) => {
            if (isChild) {
                const parentIssueRow = getParentIssueRow(issueIndex);
                if (parentIssueRow != null) {
                    const titleA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(parentIssueRow, titleColumn));
                    const childIssueA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, childIssueColumn));
                    return `=${titleA1Notation} & " - " & ${childIssueA1Notation}`;
                }
                const childIssueA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, childIssueColumn));
                return `=${childIssueA1Notation}`;
            }
            const issueA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, issueColumn));
            return `=${issueA1Notation}`;
        });
        addFormulas(estimateColumn, (row, isBuffer) => {
            if (isBuffer) {
                const startA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, startColumn));
                const endA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, endColumn));
                return `=LET(
                    workDays,
                    NETWORKDAYS(${startA1Notation}, ${endA1Notation}, ${GSheetProjectSettings.publicHolidaysRangeName}),
                    IF(
                        workDays > 0,
                        workDays,
                        ""
                    )
                )`;
            }
            return undefined;
        });
        addFormulas(startColumn, (row, isBuffer) => {
            const teamTitleA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(GSheetProjectSettings.titleRow, teamColumn));
            const estimateTitleA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(GSheetProjectSettings.titleRow, estimateColumn));
            const endTitleA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(GSheetProjectSettings.titleRow, endColumn));
            const teamA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, teamColumn));
            const earliestStartA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, earliestStartColumn));
            const deadlineA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, deadlineColumn));
            const notEnoughPreviousLanes = `
                COUNTIFS(
                    OFFSET(
                        ${teamTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    "=" & ${teamA1Notation},
                    OFFSET(
                        ${estimateTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    ">0"
                ) < resources
            `;
            const filter = `
                FILTER(
                    OFFSET(
                        ${endTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    OFFSET(
                        ${teamTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ) = ${teamA1Notation},
                    OFFSET(
                        ${estimateTitleA1Notation},
                        ${GSheetProjectSettings.firstDataRow - GSheetProjectSettings.titleRow},
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ) > 0
                )
            `;
            const lastEnd = `
                MIN(
                    SORTN(
                        ${filter},
                        resources,
                        0,
                        1,
                        FALSE
                    )
                )
            `;
            const nextWorkdayLastEnd = `
                WORKDAY(
                    ${lastEnd},
                    1,
                    ${GSheetProjectSettings.publicHolidaysRangeName}
                )
            `;
            const firstDataRowIf = `
                IF(
                    OR(
                        ROW() <= ${GSheetProjectSettings.firstDataRow},
                        ${notEnoughPreviousLanes}
                    ),
                    ${GSheetProjectSettings.settingsScheduleStartRangeName},
                    MAX(${nextWorkdayLastEnd}, ${GSheetProjectSettings.settingsScheduleStartRangeName})
                )
            `;
            let mainCalculation = `
                LET(
                    start,
                    ${firstDataRowIf},
                    IF(
                        ${earliestStartA1Notation} <> "",
                        MAX(start, ${earliestStartA1Notation}),
                        start
                    )
                )
            `;
            if (isBuffer) {
                let previousMilestone = `
                    MAX(FILTER(
                        ${GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName},
                        ${GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName} < ${deadlineA1Notation}
                    ))
                `;
                previousMilestone = `
                    MAX(
                        ${previousMilestone},
                        ${GSheetProjectSettings.settingsScheduleStartRangeName} - 1
                    )
                `;
                mainCalculation = `
                    MAX(
                        WORKDAY(
                            ${previousMilestone},
                            1,
                            ${GSheetProjectSettings.publicHolidaysRangeName}
                        ),
                        ${firstDataRowIf}
                    )
                `;
                mainCalculation = `
                    LET(
                        startDate,
                        ${mainCalculation},
                        IF(
                            startDate <= ${deadlineA1Notation},
                            startDate,
                            ""
                        )
                    )
                `;
            }
            const withResources = `
                LET(
                    resources,
                    VLOOKUP(
                        ${teamA1Notation},
                        ${GSheetProjectSettings.settingsTeamsTableRangeName},
                        1
                            + COLUMN(${GSheetProjectSettings.settingsTeamsTableResourcesRangeName})
                            - COLUMN(${GSheetProjectSettings.settingsTeamsTableRangeName}),
                        FALSE
                    ),
                    IF(
                        resources,
                        ${mainCalculation},
                        ""
                    )
                )
            `;
            const notEnoughDataIf = `
                IF(
                    ${teamA1Notation} = "",
                    "",
                    ${withResources}
                )
            `;
            return `=${notEnoughDataIf}`;
        });
        addFormulas(endColumn, (row, isBuffer) => {
            if (isBuffer) {
                const startA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, startColumn));
                const deadlineA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, deadlineColumn));
                return `=IF(
                    ${startA1Notation} <= ${deadlineA1Notation},
                    ${deadlineA1Notation},
                    ""
                )`;
            }
            const startA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, startColumn));
            const estimateA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, estimateColumn));
            const bufferRangeName = GSheetProjectSettings.settingsScheduleBufferRangeName;
            return `=
                IF(
                    OR(
                        ${startA1Notation} = "",
                        ${estimateA1Notation} = "",
                        NOT(ISNUMBER(${estimateA1Notation})),
                        ${estimateA1Notation} <= 0
                    ),
                    "",
                    WORKDAY(
                        ${startA1Notation},
                        MAX(ROUND(${estimateA1Notation} * (1 + ${bufferRangeName})) - 1, 0),
                        ${GSheetProjectSettings.publicHolidaysRangeName}
                    )
                )
            `;
        });
        addFormulas(deadlineColumn, (row, isBuffer, isChild, issueIndex) => {
            if (isChild) {
                const parentIssueRow = getParentIssueRow(issueIndex);
                if (parentIssueRow != null) {
                    const deadlineA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(parentIssueRow, deadlineColumn));
                    return `=${deadlineA1Notation}`;
                }
            }
            const milestoneA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, milestoneColumn));
            return `=
                IF(
                    ${milestoneA1Notation} = "",
                    "",
                    VLOOKUP(
                        ${milestoneA1Notation},
                        ${GSheetProjectSettings.settingsMilestonesTableRangeName},
                        1
                            + COLUMN(${GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName})
                            - COLUMN(${GSheetProjectSettings.settingsMilestonesTableRangeName}),
                        FALSE
                    )
                )
            `;
        });
    }
}
DefaultFormulas._DEFAULT_FORMULA_MARKER = "default";
DefaultFormulas._DEFAULT_CHILD_FORMULA_MARKER = "default-child";
DefaultFormulas._DEFAULT_BUFFER_FORMULA_MARKER = "default-buffer";
class DocumentFlags {
    static set(key, value = true) {
        key = `flag|${key}`;
        if (value) {
            PropertiesService.getDocumentProperties().setProperty(key, Date.now().toString());
        }
        else {
            PropertiesService.getDocumentProperties().deleteProperty(key);
        }
    }
    static isSet(key) {
        key = `flag|${key}`;
        return PropertiesService.getDocumentProperties().getProperty(key)?.length;
    }
    static cleanupByPrefix(keyPrefix) {
        keyPrefix = `flag|${keyPrefix}`;
        const entries = [];
        for (const [key, value] of Object.entries(PropertiesService.getDocumentProperties().getProperties())) {
            if (key.startsWith(keyPrefix)) {
                const number = parseFloat(value);
                if (isNaN(number)) {
                    console.warn(`Removing NaN document flag: ${key}`);
                    PropertiesService.getDocumentProperties().deleteProperty(key);
                    continue;
                }
                entries.push({ key, number });
            }
        }
        // sort ascending:
        entries.sort((e1, e2) => e1.number - e2.number);
        // skip last element:
        entries.pop();
        // remove old keys:
        for (const entry of entries) {
            PropertiesService.getDocumentProperties().deleteProperty(entry.key);
        }
    }
}

class EntryPoint {
    static entryPoint(action, useLocks) {
        if (this._isInEntryPoint) {
            return action();
        }
        let lock = null;
        if (useLocks ?? GSheetProjectSettings.useLockService) {
            lock = LockService.getDocumentLock();
        }
        try {
            this._isInEntryPoint = true;
            ExecutionCache.resetCache();
            lock?.waitLock(GSheetProjectSettings.lockTimeoutMillis);
            return action();
        }
        catch (e) {
            Observability.reportError(e);
            throw e;
        }
        finally {
            ProtectionLocks.release();
            lock?.releaseLock();
            this._isInEntryPoint = false;
        }
    }
}
EntryPoint._isInEntryPoint = false;
class ExecutionCache {
    static getOrCompute(key, compute, timerLabel) {
        const stringKey = this._getStringKey(key);
        if (this._data.has(stringKey)) {
            return this._data.get(stringKey);
        }
        if (timerLabel?.length) {
            console.time(timerLabel);
        }
        let result;
        try {
            result = compute();
        }
        finally {
            if (timerLabel?.length) {
                console.timeEnd(timerLabel);
            }
        }
        this._data.set(stringKey, result);
        return result;
    }
    static put(key, value) {
        const stringKey = this._getStringKey(key);
        this._data.set(stringKey, value);
    }
    static _getStringKey(key) {
        return JSON.stringify(key, (_, value) => {
            if (Utils.isFunction(value?.getUniqueId)) {
                return value.getUniqueId();
            }
            else if (Utils.isFunction(value?.getSheetId)) {
                return value.getSheetId();
            }
            else if (Utils.isFunction(value?.getId)) {
                return value.getId();
            }
            return value;
        });
    }
    static resetCache() {
        this._data.clear();
    }
}
ExecutionCache._data = new Map();
class Formulas {
    static processFormula(formula) {
        formula = formula.replaceAll(/#SELF_COLUMN\(([^)]+)\)/g, 'INDIRECT("RC"&COLUMN($1), FALSE)');
        formula = formula.replaceAll(/#SELF(\b|&)/g, 'INDIRECT("RC", FALSE)$1');
        return formula.split(/[\r\n]+/)
            .map(line => line.replace(/^\s+/, ''))
            .filter(line => line.length)
            .map(line => line.replaceAll(/^([<>&=*/+-]+ )/g, ' $1'))
            .map(line => line.replaceAll(/\s*\t\s*/g, ' '))
            .map(line => line.replaceAll(/"\s*&\s*""/g, '"'))
            .map(line => line.replaceAll(/([")])\s*&\s*([")])/g, '$1 & $2'))
            .map(line => line + (line.endsWith(',') || line.endsWith(';') ? ' ' : ''))
            .join('')
            .trim();
    }
    static addFormulaMarker(formula, marker) {
        if (!marker?.length) {
            return formula;
        }
        formula = formula.replace(/^=+/, '');
        formula = `IF("GSPf"<>"${marker}", ${formula}, "")`;
        return '=' + formula;
    }
    static addFormulaMarkers(formula, ...markers) {
        markers = markers.filter(it => it?.length);
        if (!markers?.length) {
            return formula;
        }
        if (markers.length === 1) {
            return this.addFormulaMarker(formula, markers[0]);
        }
        formula = formula.replace(/^=+/, '');
        formula = `IF(AND(${markers.map(marker => `"GSPf"<>"${marker}"`).join(', ')}), ${formula}, "")`;
        return '=' + formula;
    }
    static extractFormulaMarkers(formula) {
        if (!formula?.length) {
            return [];
        }
        const markers = Utils.arrayOf();
        const regex = /"GSPf"\s*<>\s*"([^"]+)"/g;
        let match;
        while ((match = regex.exec(formula)) !== null) {
            markers.push(match[1]);
        }
        return markers;
    }
    static escapeFormulaString(string) {
        return string.replaceAll(/"/g, '""');
    }
}
class Images {
}
Images.loadingImageUrl = 'https://raw.githubusercontent.com/remal/misc/main/spinner-100.gif';
class IssueDataDisplay extends AbstractIssueLogic {
    static reloadIssueData(range) {
        const processedRange = this._processRange(range);
        if (processedRange == null) {
            return;
        }
        else {
            range = processedRange;
        }
        const sheet = range.getSheet();
        const iconColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.iconColumnName);
        const issueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName);
        const childIssueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueKeyColumnName);
        const lastDataReloadColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.lastDataReloadColumnName);
        const titleColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.titleColumnName);
        const { issues, childIssues, lastDataReload } = this._getIssueValuesWithLastReloadDate(range);
        let lastDataNotChangedCheckTimestamp = Date.now();
        const indexes = Array.from(Utils.range(0, issues.length - 1))
            .toSorted((i1, i2) => {
            const d1 = lastDataReload[i1];
            const d2 = lastDataReload[i2];
            if (d1 == null && d2 == null) {
                return 0;
            }
            else if (d1 != null && d2 != null) {
                return d1.getTime() - d2.getTime();
            }
            else if (d1 != null) {
                return 1;
            }
            else {
                return -1;
            }
        });
        const processIndex = (index) => {
            const row = range.getRow() + index;
            ProtectionLocks.lockRows(sheet, row);
            const cleanupColumns = (withTitle = false) => {
                const notations = [
                    [
                        withTitle ? titleColumn : null,
                        iconColumn,
                    ],
                    [
                        GSheetProjectSettings.issuesMetrics,
                        GSheetProjectSettings.counterIssuesMetrics,
                    ]
                        .flatMap(metrics => Object.keys(metrics))
                        .map(columnName => SheetUtils.findColumnByName(sheet, columnName)),
                ]
                    .flat()
                    .filter(column => column != null)
                    .map(column => sheet.getRange(row, column))
                    .map(range => range.getA1Notation())
                    .filter(Utils.distinct());
                if (notations.length) {
                    sheet.getRangeList(notations).setValue('');
                }
                sheet.getRange(row, lastDataReloadColumn).setValue(new Date());
            };
            if (GSheetProjectSettings.skipHiddenIssues && sheet.isRowHiddenByUser(row)) { // a slow check
                cleanupColumns();
                return;
            }
            let currentIssueColumn;
            let originalIssueKeysText;
            let isChildIssue = false;
            if (childIssues[index]?.length) {
                currentIssueColumn = childIssueColumn;
                originalIssueKeysText = childIssues[index];
                isChildIssue = true;
            }
            else if (issues[index]?.length) {
                currentIssueColumn = issueColumn;
                originalIssueKeysText = issues[index];
            }
            else {
                cleanupColumns(true);
                return;
            }
            if (GSheetProjectSettings.notIssueKeyRegex?.test(originalIssueKeysText)) {
                cleanupColumns(true);
                return;
            }
            const originalIssueKeysRange = sheet.getRange(row, currentIssueColumn);
            const isOriginalIssueKeysTextChanged = () => {
                const now = Date.now();
                const minTimestamp = now - GSheetProjectSettings.originalIssueKeysTextChangedTimeout;
                if (lastDataNotChangedCheckTimestamp >= minTimestamp) {
                    return false;
                }
                const currentValue = originalIssueKeysRange.getValue().toString();
                lastDataNotChangedCheckTimestamp = now;
                if (currentValue !== originalIssueKeysText) {
                    Observability.reportWarning(`Content of ${originalIssueKeysRange.getA1Notation()} has been changed`);
                    return true;
                }
                return false;
            };
            const allIssueKeys = originalIssueKeysText
                .split(/[\r\n]+/)
                .map(key => key.trim())
                .filter(key => key.length)
                .filter(Utils.distinct());
            let issueTracker = null;
            const issueKeys = Utils.arrayOf();
            const issueKeyIds = {};
            const issueKeyQueries = {};
            for (let issueKey of allIssueKeys) {
                if (issueTracker != null) {
                    if (!issueTracker.supportsIssueKey(issueKey)) {
                        continue;
                    }
                }
                else {
                    const keyTracker = GSheetProjectSettings.issueTrackers.find(it => it.supportsIssueKey(issueKey));
                    if (keyTracker != null) {
                        issueTracker = keyTracker;
                    }
                    else {
                        continue;
                    }
                }
                issueKeys.push(issueKey);
                const issueId = issueTracker.extractIssueId(issueKey);
                if (issueId?.length) {
                    issueKeyIds[issueKey] = issueId;
                }
                const searchQuery = issueTracker.extractSearchQuery(issueKey);
                if (searchQuery?.length) {
                    issueKeyQueries[issueKey] = searchQuery;
                }
            }
            if (issueTracker == null) {
                cleanupColumns();
                return;
            }
            const allIssueLinks = allIssueKeys.map(issueKey => {
                if (issueKeys.includes(issueKey)) {
                    const issueId = issueKeyIds[issueKey];
                    if (issueId?.length) {
                        return {
                            title: issueTracker.canonizeIssueKey(issueKey),
                            url: issueTracker.getUrlForIssueId(issueId),
                        };
                    }
                    else {
                        const searchQuery = issueKeyQueries[issueKey];
                        if (searchQuery?.length) {
                            return {
                                title: issueTracker.canonizeIssueKey(issueKey),
                                url: issueTracker.getUrlForSearchQuery(searchQuery),
                            };
                        }
                    }
                }
                return {
                    title: issueKey,
                };
            });
            if (isOriginalIssueKeysTextChanged()) {
                return;
            }
            const issuesRichTextValue = RichTextUtils.createLinksValue(allIssueLinks);
            originalIssueKeysText = issuesRichTextValue.getText();
            sheet.getRange(row, currentIssueColumn).setRichTextValue(issuesRichTextValue);
            const loadedIssues = LazyProxy.create(() => Observability.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading issues`,
            ].join(': '), () => {
                const issueIds = Object.values(issueKeyIds).filter(Utils.distinct());
                return issueTracker?.loadIssuesByIssueId(issueIds);
            }));
            const loadedChildIssues = LazyProxy.create(() => Observability.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading child issues`,
            ].join(': '), () => {
                const issueIds = loadedIssues.map(it => it.id);
                return [
                    issueTracker.loadChildrenFor(loadedIssues),
                    Object.values(issueKeyQueries)
                        .filter(Utils.distinct())
                        .flatMap(query => issueTracker.searchByQuery(query)),
                ]
                    .flat()
                    .filter(Utils.distinctBy(issue => issue.id))
                    .filter(issue => !issueIds.includes(issue.id));
            }));
            const loadedBlockerIssues = LazyProxy.create(() => Observability.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading blocker issues`,
            ].join(': '), () => {
                const issueIds = loadedIssues.map(it => it.id);
                const allIssues = loadedIssues.concat(loadedChildIssues);
                return issueTracker.loadBlockersFor(allIssues)
                    .filter(issue => !issueIds.includes(issue.id));
            }));
            const titles = issueKeys.map(issueKey => {
                const issueId = issueKeyIds[issueKey];
                if (issueId?.length) {
                    return loadedIssues.find(issue => issue.id === issueId)?.title;
                }
                if (issueKeyQueries[issueKey]?.length) {
                    return Observability.timed([
                        IssueDataDisplay.name,
                        this.reloadIssueData.name,
                        `row #${row}`,
                        `loading search title for "${issueKey}" issue key`,
                    ].join(': '), () => issueTracker.loadIssueKeySearchTitle(issueKey));
                }
                return undefined;
            })
                .map(title => title?.trim())
                .filter(title => title?.length)
                .map(title => title);
            if (isOriginalIssueKeysTextChanged()) {
                return;
            }
            sheet.getRange(row, titleColumn).setValue(titles.join('\n'));
            for (const handler of GSheetProjectSettings.onIssuesLoadedHandlers) {
                if (isOriginalIssueKeysTextChanged()) {
                    return;
                }
                handler(loadedIssues, sheet, row, isChildIssue);
            }
            for (const [columnName, issuesMetric] of Object.entries(GSheetProjectSettings.issuesMetrics)) {
                const column = SheetUtils.findColumnByName(sheet, columnName);
                if (column == null) {
                    continue;
                }
                let value = issuesMetric(loadedIssues, loadedChildIssues, loadedBlockerIssues, sheet, row);
                if (value == null) {
                    value = '';
                }
                else if (Utils.isBoolean(value)) {
                    value = value ? "Yes" : "";
                }
                if (isOriginalIssueKeysTextChanged()) {
                    return;
                }
                sheet.getRange(row, column).setValue(value);
            }
            for (const [columnName, issuesMetric] of Object.entries(GSheetProjectSettings.counterIssuesMetrics)) {
                const column = SheetUtils.findColumnByName(sheet, columnName);
                if (column == null) {
                    continue;
                }
                const foundIssues = issuesMetric(loadedIssues, loadedChildIssues, loadedBlockerIssues, sheet, row).filter(Utils.distinctBy(issue => issue.id));
                if (!foundIssues.length) {
                    sheet.getRange(row, column).setValue('');
                    continue;
                }
                const foundIssueIds = foundIssues.map(issue => issue.id);
                const link = {
                    title: foundIssues.length.toString(),
                    url: issueTracker.getUrlForIssueIds(foundIssueIds),
                };
                if (isOriginalIssueKeysTextChanged()) {
                    return;
                }
                sheet.getRange(row, column).setRichTextValue(RichTextUtils.createLinkValue(link));
            }
            if (isOriginalIssueKeysTextChanged()) {
                return;
            }
            sheet.getRange(row, lastDataReloadColumn).setValue(allIssueKeys.length ? new Date() : '');
        };
        const start = Date.now();
        let processedIndexes = 0;
        for (const index of indexes) {
            const row = range.getRow() + index;
            console.info(`Processing index ${index} (${++processedIndexes} / ${indexes.length}), row #${row}`);
            if (Date.now() - start >= GSheetProjectSettings.issuesLoadTimeoutMillis) {
                Observability.reportWarning("Issues load timeout occurred");
                break;
            }
            const iconRange = sheet.getRange(row, iconColumn);
            try {
                Observability.timed(`loading issue data for row #${row}`, () => {
                    if (GSheetProjectSettings.loadingText?.length) {
                        iconRange.setValue(GSheetProjectSettings.loadingText);
                    }
                    else {
                        iconRange.setFormula(`=IMAGE("${Images.loadingImageUrl}")`);
                    }
                    SpreadsheetApp.flush();
                    processIndex(index);
                });
            }
            catch (e) {
                Observability.reportError(`Error loading issue data for row #${row}: ${e}`, e);
            }
            finally {
                iconRange.setValue('');
                SpreadsheetApp.flush();
            }
        }
    }
    static reloadAllIssuesData() {
        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName);
        const range = sheet.getRange(1, 1, SheetUtils.getLastRow(sheet), SheetUtils.getLastColumn(sheet));
        this.reloadIssueData(range);
    }
}
class IssueHierarchyFormatter extends AbstractIssueLogic {
    static formatHierarchy(range) {
        const processedRange = this._processRange(range);
        if (processedRange == null) {
            return;
        }
        else {
            range = processedRange;
        }
        const sheet = range.getSheet();
        const startRow = range.getRow();
        const endRow = startRow + range.getNumRows() - 1;
        const { issues, childIssues } = this._getIssueValues(sheet.getRange(GSheetProjectSettings.firstDataRow, range.getColumn(), endRow - GSheetProjectSettings.firstDataRow + 1, range.getNumColumns()));
        const issueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueKeyColumnName);
        for (let row = startRow; row <= endRow; ++row) {
            const index = row - GSheetProjectSettings.firstDataRow;
            const issue = issues[index];
            const childIssue = childIssues[index];
            if (!issue?.length) {
                continue;
            }
            const issueRange = sheet.getRange(row, issueColumn);
            if (!childIssue?.length) {
                issueRange.setFontSize(GSheetProjectSettings.fontSize);
                continue;
            }
            const parentIssueIndex = issues.indexOf(issue);
            if (parentIssueIndex < 0) {
                continue;
            }
            if (childIssues[parentIssueIndex]?.length) {
                continue;
            }
            const parentIssueRow = GSheetProjectSettings.firstDataRow + parentIssueIndex;
            const parentIssueRange = sheet.getRange(parentIssueRow, issueColumn);
            issueRange
                .setFormula(Formulas.processFormula(`=
                    ${RangeUtils.getAbsoluteA1Notation(parentIssueRange)}
                `))
                .setFontSize(GSheetProjectSettings.fontSize - 2);
        }
    }
}
class IssueTracker {
    supportsIssueKey(issueKey) {
        return this.extractIssueId(issueKey) != null
            || this.extractSearchQuery(issueKey) != null;
    }
    canonizeIssueKey(issueKey) {
        {
            const issueId = this.extractIssueId(issueKey);
            if (issueId?.length) {
                const canonizedKey = this.issueIdToIssueKey(issueId);
                if (canonizedKey?.length) {
                    return canonizedKey;
                }
            }
        }
        {
            const searchQuery = this.extractSearchQuery(issueKey);
            if (searchQuery?.length) {
                const canonizedKey = this.searchQueryToIssueKey(searchQuery);
                if (canonizedKey?.length) {
                    return canonizedKey;
                }
            }
        }
        return issueKey;
    }
    extractIssueId(issueKey) {
        throw Utils.throwNotImplemented(this.constructor.name, this.extractIssueId.name);
    }
    issueIdToIssueKey(issueId) {
        throw Utils.throwNotImplemented(this.constructor.name, this.issueIdToIssueKey.name);
    }
    getUrlForIssueId(issueId) {
        return this.getUrlForIssueIds([issueId]);
    }
    getUrlForIssueIds(issueIds) {
        if (!issueIds?.length) {
            return undefined;
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.getUrlForIssueIds.name);
    }
    loadIssuesByIssueId(issueIds) {
        if (!issueIds?.length) {
            return [];
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.loadIssuesByIssueId.name);
    }
    loadChildrenFor(issues) {
        if (!issues?.length) {
            return [];
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.loadChildrenFor.name);
    }
    loadBlockersFor(issues) {
        if (!issues?.length) {
            return [];
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.loadBlockersFor.name);
    }
    extractSearchQuery(issueKey) {
        throw Utils.throwNotImplemented(this.constructor.name, this.extractSearchQuery.name);
    }
    searchQueryToIssueKey(query) {
        throw Utils.throwNotImplemented(this.constructor.name, this.issueIdToIssueKey.name);
    }
    getUrlForSearchQuery(query) {
        throw Utils.throwNotImplemented(this.constructor.name, this.getUrlForSearchQuery.name);
    }
    loadIssueKeySearchTitle(issueKey) {
        return this.extractSearchQuery(issueKey);
    }
    searchByQuery(query) {
        if (!query?.length) {
            return [];
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.searchByQuery.name);
    }
}
class Issue {
    constructor(issueTracker) {
        this.issueTracker = issueTracker;
    }
    get id() {
        throw Utils.throwNotImplemented(this.constructor.name, 'id');
    }
    get title() {
        throw Utils.throwNotImplemented(this.constructor.name, 'title');
    }
    get type() {
        throw Utils.throwNotImplemented(this.constructor.name, 'type');
    }
    get status() {
        throw Utils.throwNotImplemented(this.constructor.name, 'status');
    }
    get open() {
        throw Utils.throwNotImplemented(this.constructor.name, 'open');
    }
    get assignee() {
        throw Utils.throwNotImplemented(this.constructor.name, 'assignee');
    }
}
class IssueTrackerExample extends IssueTracker {
    extractIssueId(issueKey) {
        return Utils.extractRegex(issueKey, /^example\/([\d.-]+)$/, 1);
    }
    issueIdToIssueKey(issueId) {
        return `example/${issueId}`;
    }
    getUrlForIssueId(issueId) {
        return `https://example.com/issues/${encodeURIComponent(issueId)}`;
    }
    getUrlForIssueIds(issueIds) {
        if (!issueIds?.length) {
            return null;
        }
        return `https://example.com/search?q=id:(${encodeURIComponent(issueIds.join('|'))})`;
    }
    loadIssuesByIssueId(issueIds) {
        if (!issueIds?.length) {
            return [];
        }
        return issueIds.map(id => new IssueExample(this, id));
    }
    loadChildrenFor(issues) {
        return issues.map(issue => issue.id).flatMap(id => {
            let hash = parseInt(id);
            if (isNaN(hash)) {
                hash = Math.abs(Utils.hashCode(id));
            }
            return Array.from(Utils.range(0, hash % 3)).map(index => new IssueExample(this, `${id}-${index + 1}`));
        });
    }
    loadBlockersFor(issues) {
        return issues.map(issue => issue.id).flatMap(id => {
            let hash = parseInt(id);
            if (isNaN(hash)) {
                hash = Math.abs(Utils.hashCode(id));
            }
            return Array.from(Utils.range(0, hash % 2)).map(index => new IssueExample(this, `${id}-blocker-${index + 1}`));
        });
    }
    extractSearchQuery(issueKey) {
        return Utils.extractRegex(issueKey, /^example\/search\/(.+)$/, 1);
    }
    searchQueryToIssueKey(query) {
        return `example/search/${query}`;
    }
    getUrlForSearchQuery(query) {
        return `https://example.com/search?q=${encodeURIComponent(query)}`;
    }
    searchByQuery(query) {
        if (!query?.length) {
            return [];
        }
        const hash = Math.abs(Utils.hashCode(query));
        return Array.from(Utils.range(0, hash % 3)).map(index => new IssueExample(this, `search-${hash}-${index + 1}`));
    }
}
GSheetProjectSettings.issueTrackers.push(new IssueTrackerExample());
class IssueExample extends Issue {
    constructor(issueTracker, id) {
        super(issueTracker);
        this._id = id;
    }
    get id() {
        return this._id;
    }
    get title() {
        return `Issue '${this.id}'`;
    }
    get type() {
        return 'task';
    }
    get status() {
        let hash = parseInt(this.id);
        if (isNaN(hash)) {
            hash = Math.abs(Utils.hashCode(this.id));
        }
        return hash % 3 !== 0
            ? 'open'
            : 'closed';
    }
    get open() {
        return this.status === 'open';
    }
    get assignee() {
        let hash = parseInt(this.id);
        if (isNaN(hash)) {
            hash = Math.abs(Utils.hashCode(this.id));
        }
        return hash.toString();
    }
}
class Lazy {
    constructor(supplier) {
        this._supplier = supplier;
    }
    get() {
        if (this._supplier != null) {
            this._value = this._supplier();
            this._supplier = undefined;
        }
        return this._value;
    }
}
class LazyProxy {
    static create(supplier) {
        const lazy = new Lazy(supplier);
        const proxy = new Proxy(lazy, {
            apply(lazy, thisArg, argArray) {
                const instance = lazy.get();
                argArray = argArray.map(it => this.unwrapLazyProxy(it));
                return Reflect.apply(instance, thisArg, argArray);
            },
            construct(lazy, argArray, newTarget) {
                const instance = lazy.get();
                argArray = argArray.map(it => this.unwrapLazyProxy(it));
                return Reflect.construct(instance, argArray, newTarget);
            },
            defineProperty(lazy, property, attributes) {
                const instance = lazy.get();
                return Reflect.defineProperty(instance, property, attributes);
            },
            deleteProperty(lazy, property) {
                const instance = lazy.get();
                return Reflect.deleteProperty(instance, property);
            },
            get(lazy, property, receiver) {
                const instance = lazy.get();
                let value = Reflect.get(instance, property, instance);
                if (Utils.isFunction(value)) {
                    return function () {
                        const target = this === receiver ? instance : this;
                        const argArray = Array.from(arguments).map(it => LazyProxy.unwrapLazyProxy(it));
                        return value.apply(target, argArray);
                    };
                }
                return value;
            },
            getOwnPropertyDescriptor(lazy, property) {
                const instance = lazy.get();
                return Reflect.getOwnPropertyDescriptor(instance, property);
            },
            getPrototypeOf(lazy) {
                const instance = lazy.get();
                return Reflect.getPrototypeOf(instance);
            },
            has(lazy, property) {
                const instance = lazy.get();
                return Reflect.has(instance, property);
            },
            isExtensible(lazy) {
                const instance = lazy.get();
                return Reflect.isExtensible(instance);
            },
            ownKeys(lazy) {
                const instance = lazy.get();
                return Reflect.ownKeys(instance);
            },
            preventExtensions(lazy) {
                const instance = lazy.get();
                return Reflect.preventExtensions(instance);
            },
            set(lazy, property, newValue) {
                const instance = lazy.get();
                return Reflect.set(instance, property, newValue, instance);
            },
            setPrototypeOf(lazy, value) {
                const instance = lazy.get();
                return Reflect.setPrototypeOf(instance, value);
            },
        });
        this._lazyProxyToLazy.set(proxy, lazy);
        return proxy;
    }
    static unwrapLazyProxy(lazyProxy) {
        const lazy = this._lazyProxyToLazy.get(lazyProxy);
        if (lazy == null) {
            return lazyProxy;
        }
        return lazy.get();
    }
}
LazyProxy._lazyProxyToLazy = new WeakMap();
class NamedRangeUtils {
    static findNamedRange(rangeName) {
        const namedRanges = ExecutionCache.getOrCompute('named-ranges', () => {
            const result = new Map();
            for (const namedRange of SpreadsheetApp.getActiveSpreadsheet().getNamedRanges()) {
                const name = Utils.normalizeName(namedRange.getName());
                result.set(name, namedRange);
            }
            return result;
        }, `${NamedRangeUtils.name}: ${this.findNamedRange.name}`);
        rangeName = Utils.normalizeName(rangeName);
        return namedRanges.get(rangeName);
    }
    static getNamedRange(rangeName) {
        return this.findNamedRange(rangeName) ?? (() => {
            throw new Error(`"${rangeName}" named range can't be found`);
        })();
    }
    static getNamedRangeColumn(rangeName) {
        return this.getNamedRange(rangeName).getRange().getColumn();
    }
}
class Observability {
    static reportError(message, exception) {
        console.error(message);
        SpreadsheetApp.getActiveSpreadsheet().toast(message?.toString() ?? '', "Automation error");
        if (exception != null) {
            console.log(exception);
        }
    }
    static reportWarning(message) {
        console.warn(message);
    }
    static timed(timerLabel, action, enabled) {
        if (enabled === false) {
            return action();
        }
        console.time(timerLabel);
        try {
            return action();
        }
        finally {
            console.timeEnd(timerLabel);
        }
    }
}
class PropertyLocks {
    static waitLock(property, timeout = GSheetProjectSettings.lockTimeoutMillis) {
        property = `lock|${property}`;
        const start = Date.now();
        while (true) {
            const propertyValue = PropertiesService.getDocumentProperties().getProperty(property);
            if (!propertyValue?.length) {
                break;
            }
            const date = Utils.parseDate(propertyValue);
            if (date == null || date.getTime() < Date.now()) {
                break;
            }
            if (start + timeout > Date.now()) {
                Utilities.sleep(1000);
            }
            else {
                return false;
            }
        }
        PropertiesService.getDocumentProperties().setProperty(property, (Date.now() + timeout).toString());
        return true;
    }
    static releaseLock(property) {
        property = `lock|${property}`;
        PropertiesService.getDocumentProperties().deleteProperty(property);
    }
    static releaseExpiredPropertyLocks() {
        for (const [key, value] of Object.entries(PropertiesService.getDocumentProperties().getProperties())) {
            if (!key.startsWith('lock|')) {
                continue;
            }
            const date = Utils.parseDate(value);
            if (date == null || date.getTime() < Date.now()) {
                PropertiesService.getDocumentProperties().deleteProperty(key);
            }
        }
    }
}
class ProtectionLocks {
    static lockAllColumns(sheet) {
        if (!GSheetProjectSettings.lockColumns) {
            return;
        }
        const sheetId = sheet.getSheetId();
        if (this._allColumnsProtections.has(sheetId)) {
            return;
        }
        Observability.timed(`${ProtectionLocks.name}: ${this.lockAllColumns.name}: ${sheet.getSheetName()}`, () => {
            const range = sheet.getRange(1, 1, 1, SheetUtils.getMaxColumns(sheet));
            const protection = range.protect()
                .setDescription(`lock|columns|all|${Date.now()}`)
                .setWarningOnly(true);
            this._allColumnsProtections.set(sheetId, protection);
        });
    }
    static lockAllRows(sheet) {
        if (!GSheetProjectSettings.lockRows) {
            return;
        }
        const sheetId = sheet.getSheetId();
        if (this._allRowsProtections.has(sheetId)) {
            return;
        }
        Observability.timed(`${ProtectionLocks.name}: ${this.lockAllRows.name}: ${sheet.getSheetName()}`, () => {
            const range = sheet.getRange(1, SheetUtils.getMaxColumns(sheet), SheetUtils.getMaxRows(sheet), 1);
            const protection = range.protect()
                .setDescription(`lock|rows|all|${Date.now()}`)
                .setWarningOnly(true);
            this._allRowsProtections.set(sheetId, protection);
        });
    }
    static lockRows(sheet, rowToLock) {
        if (!GSheetProjectSettings.lockRows) {
            return;
        }
        if (rowToLock <= 0) {
            return;
        }
        const sheetId = sheet.getSheetId();
        if (this._allRowsProtections.has(sheetId)) {
            return;
        }
        if (!this._rowsProtections.has(sheetId)) {
            this._rowsProtections.set(sheetId, new Map());
        }
        const rowsProtections = this._rowsProtections.get(sheetId);
        const maxLockedRow = Array.from(rowsProtections.keys()).reduce((prev, cur) => Math.max(prev, cur), 0);
        if (maxLockedRow < rowToLock) {
            Observability.timed(`${ProtectionLocks.name}: ${this.lockRows.name}: ${sheet.getSheetName()}: ${rowToLock}`, () => {
                const range = sheet.getRange(1, SheetUtils.getMaxColumns(sheet), rowToLock, 1);
                const protection = range.protect()
                    .setDescription(`lock|rows|${rowToLock}|${Date.now()}`)
                    .setWarningOnly(true);
                rowsProtections.set(rowToLock, protection);
            });
        }
    }
    static release() {
        if (!GSheetProjectSettings.lockColumns && !GSheetProjectSettings.lockRows) {
            return;
        }
        Observability.timed(`${ProtectionLocks.name}: ${this.release.name}`, () => {
            this._allColumnsProtections.forEach(protection => protection.remove());
            this._allColumnsProtections.clear();
            this._allRowsProtections.forEach(protection => protection.remove());
            this._allRowsProtections.clear();
            this._rowsProtections.forEach(protections => Array.from(protections.values()).forEach(protection => protection.remove()));
            this._rowsProtections.clear();
        });
    }
    static releaseExpiredLocks() {
        Observability.timed(`${ProtectionLocks.name}: ${this.releaseExpiredLocks.name}`, () => {
            const maxLockDurationMillis = 10 * 60 * 1000;
            const minTimestamp = Date.now() - maxLockDurationMillis;
            SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => {
                for (const protection of sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)) {
                    const description = protection.getDescription();
                    if (!description.startsWith('lock|')) {
                        continue;
                    }
                    const date = Utils.parseDate(description.split('|').slice(-1)[0]);
                    if (date != null && date.getTime() < minTimestamp) {
                        console.warn(`Removing expired protection lock: ${description}`);
                        protection.remove();
                    }
                }
            });
        });
    }
}
ProtectionLocks._allColumnsProtections = new Map();
ProtectionLocks._allRowsProtections = new Map();
ProtectionLocks._rowsProtections = new Map();
class RangeUtils {
    static isRangeSheet(range, sheet) {
        if (range == null) {
            return false;
        }
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.findSheetByName(sheet);
        }
        if (sheet == null) {
            return false;
        }
        return range.getSheet().getSheetId() === sheet.getSheetId();
    }
    static getAbsoluteA1Notation(range) {
        return range.getA1Notation()
            .replaceAll(/[A-Z]+/g, '$$$&')
            .replaceAll(/\d+/g, '$$$&');
    }
    static getAbsoluteReferenceFormula(range) {
        return '=' + this.getAbsoluteA1Notation(range);
    }
    static toColumnRange(range, column) {
        if (range == null || column == null) {
            return undefined;
        }
        if (Utils.isString(column)) {
            column = SheetUtils.findColumnByName(range.getSheet(), column);
        }
        if (column == null) {
            return undefined;
        }
        return range.offset(0, column - range.getColumn(), range.getNumRows(), 1);
    }
    static withMinRow(range, minRow) {
        const startRow = range.getRow();
        const rowDiff = minRow - startRow;
        if (rowDiff <= 0) {
            return range;
        }
        return range.offset(rowDiff, 0, Math.max(range.getNumRows() - rowDiff, 1), range.getNumColumns());
    }
    static withMaxRow(range, maxRow) {
        const startRow = range.getRow();
        const endRow = startRow + range.getNumRows() - 1;
        if (maxRow >= endRow) {
            return range;
        }
        return range.offset(0, 0, Math.max(maxRow - startRow + 1, 1), range.getNumColumns());
    }
    static withMinMaxRows(range, minRow, maxRow) {
        range = this.withMinRow(range, minRow);
        range = this.withMaxRow(range, maxRow);
        return range;
    }
    static doesRangeHaveColumn(range, column) {
        if (range == null) {
            return false;
        }
        if (Utils.isString(column)) {
            column = SheetUtils.findColumnByName(range.getSheet(), column);
        }
        if (column == null) {
            return false;
        }
        const minColumn = range.getColumn();
        const maxColumn = minColumn + range.getNumColumns() - 1;
        return minColumn <= column && column <= maxColumn;
    }
    static doesRangeHaveSheetColumn(range, sheet, column) {
        return this.isRangeSheet(range, sheet) && this.doesRangeHaveColumn(range, column);
    }
    static doesRangeIntersectsWithNamedRange(range, namedRange) {
        if (range == null || namedRange == null) {
            return false;
        }
        if (Utils.isString(namedRange)) {
            namedRange = NamedRangeUtils.findNamedRange(namedRange);
        }
        if (namedRange == null) {
            return false;
        }
        const rangeToFind = namedRange.getRange();
        if (range.getSheet().getSheetId() !== namedRange.getRange().getSheet().getSheetId()) {
            return false;
        }
        const minColumnToFind = rangeToFind.getColumn();
        const maxColumnToFind = minColumnToFind + rangeToFind.getNumColumns() - 1;
        const minColumn = range.getColumn();
        const maxColumn = minColumn + range.getNumColumns() - 1;
        if (maxColumnToFind < minColumn || maxColumn < minColumnToFind) {
            return false;
        }
        const minRowToFind = rangeToFind.getRow();
        const maxRowToFind = minRowToFind + rangeToFind.getNumRows() - 1;
        const minRow = range.getRow();
        const maxRow = minRow + range.getNumRows() - 1;
        if (maxRowToFind < minRow || maxRow < minRowToFind) {
            return false;
        }
        return true;
    }
    static getIndent(range) {
        const numberFormat = range.getNumberFormat();
        return this._parseIndent(numberFormat);
    }
    static setIndent(range, indent) {
        indent = Math.max(indent, 0);
        let numberFormat = range.getNumberFormat();
        if (indent === this._parseIndent(numberFormat)) {
            return;
        }
        numberFormat = numberFormat.trim();
        if (numberFormat.length) {
            range.setNumberFormat(`${' '.repeat(indent)}${numberFormat}`);
        }
        else if (indent > 0) {
            range.setNumberFormat(`${' '.repeat(indent)}@`);
        }
        else {
            // do nothing
        }
    }
    static setStringIndent(range, indent) {
        indent = Math.max(indent, 0);
        range.setNumberFormat(`${' '.repeat(indent)}@`);
    }
    static _parseIndent(numberFormat) {
        const indentMatch = numberFormat.match(/^( +)/);
        if (indentMatch) {
            return indentMatch[0].length;
        }
        return 0;
    }
}
class RichTextUtils {
    static createLinkValue(link) {
        return this.createLinksValue([link]);
    }
    static createLinksValue(links) {
        let text = '';
        const linksWithOffsets = [];
        links.forEach(link => {
            const title = link.title?.length
                ? link.title
                : link.url;
            if (!title?.length) {
                return;
            }
            if (text.length) {
                text += '\n';
            }
            if (link.url?.length) {
                linksWithOffsets.push({
                    url: link.url,
                    start: text.length,
                    end: text.length + title.length,
                });
            }
            text += title;
        });
        const builder = SpreadsheetApp.newRichTextValue().setText(text);
        linksWithOffsets.forEach(link => builder.setLinkUrl(link.start, link.end, link.url));
        builder.setTextStyle(SpreadsheetApp.newTextStyle()
            .setUnderline(false)
            .build());
        return builder.build();
    }
}
class SheetLayout {
    get sheet() {
        const sheetName = this.sheetName;
        let sheet = SheetUtils.findSheetByName(sheetName);
        if (sheet == null) {
            sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
            ExecutionCache.resetCache();
        }
        return sheet;
    }
    get _documentFlagPrefix() {
        return `${this.constructor?.name || Utils.normalizeName(this.sheetName)}:migrate:`;
    }
    get _documentFlag() {
        return `${this._documentFlagPrefix}afa7daedef473fa0a7f16e87ceb6711a694cf2a9001cc08ef03d0e60bacc6a27:${GSheetProjectSettings.computeStringSettingsHash()}`;
    }
    migrateIfNeeded() {
        if (DocumentFlags.isSet(this._documentFlag)) {
            return false;
        }
        this.migrate();
        return true;
    }
    migrate() {
        const sheet = this.sheet;
        const conditionalFormattingScope = `layout:${this.constructor?.name || Utils.normalizeName(this.sheetName)}`;
        let conditionalFormattingOrder = 0;
        ConditionalFormatting.removeConditionalFormatRulesByScope(sheet, 'layout');
        ConditionalFormatting.removeConditionalFormatRulesByScope(sheet, conditionalFormattingScope);
        const columns = this.columns.reduce((map, info) => {
            map.set(Utils.normalizeName(info.name), info);
            return map;
        }, new Map());
        if (!columns.size) {
            DocumentFlags.set(this._documentFlag);
            DocumentFlags.cleanupByPrefix(this._documentFlagPrefix);
            return;
        }
        ProtectionLocks.lockAllColumns(sheet);
        let lastColumn = SheetUtils.getLastColumn(sheet);
        const maxRows = SheetUtils.getMaxRows(sheet);
        const existingNormalizedNames = sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn)
            .getValues()[0]
            .map(it => it?.toString())
            .map(it => it?.length ? Utils.normalizeName(it) : '');
        for (const [columnName, info] of columns.entries()) {
            const existingIndex = existingNormalizedNames.indexOf(columnName);
            if (existingIndex >= 0) {
                continue;
            }
            console.info(`Adding "${info.name}" column`);
            ++lastColumn;
            const titleRange = sheet.getRange(GSheetProjectSettings.titleRow, lastColumn)
                .setValue(info.name);
            ExecutionCache.resetCache();
            if (info.defaultTitleFontSize != null && info.defaultTitleFontSize > 0) {
                titleRange.setFontSize(info.defaultTitleFontSize);
            }
            if (Utils.isNumber(info.defaultWidth)) {
                sheet.setColumnWidth(lastColumn, info.defaultWidth);
            }
            else if (info.defaultWidth === '#default-height') {
                sheet.setColumnWidth(lastColumn, 21);
            }
            else if (info.defaultWidth === '#height') {
                const height = sheet.getRowHeight(1);
                sheet.setColumnWidth(lastColumn, height);
            }
            if (info.defaultFormat != null) {
                sheet.getRange(GSheetProjectSettings.firstDataRow, lastColumn, maxRows, 1)
                    .setNumberFormat(info.defaultFormat);
            }
            if (info.defaultHorizontalAlignment?.length) {
                sheet.getRange(GSheetProjectSettings.firstDataRow, lastColumn, maxRows, 1)
                    .setHorizontalAlignment(info.defaultHorizontalAlignment);
            }
            if (info.hiddenByDefault) {
                sheet.hideColumns(lastColumn);
            }
            existingNormalizedNames.push(columnName);
        }
        SheetUtils.setLastColumn(sheet, lastColumn);
        const existingFormulas = new Lazy(() => sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn).getFormulas()[0]);
        for (const [columnName, info] of columns.entries()) {
            const index = existingNormalizedNames.indexOf(columnName);
            if (index < 0) {
                continue;
            }
            const column = index + 1;
            if (info.arrayFormula?.length) {
                const formulaToExpect = Formulas.processFormula(`=
                    {
                        "${Formulas.escapeFormulaString(info.name)}";
                        ${Formulas.processFormula(info.arrayFormula)}
                    }
                `);
                const formula = existingFormulas.get()[index];
                if (formula !== formulaToExpect) {
                    sheet.getRange(GSheetProjectSettings.titleRow, column)
                        .setFormula(formulaToExpect);
                }
            }
            const range = sheet.getRange(GSheetProjectSettings.firstDataRow, column, maxRows - GSheetProjectSettings.firstDataRow + 1, 1);
            if (info.rangeName?.length) {
                SpreadsheetApp.getActiveSpreadsheet().setNamedRange(info.rangeName, range);
            }
            let dataValidation = info.dataValidation != null
                ? info.dataValidation()
                : null;
            if (dataValidation != null) {
                if (dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CUSTOM_FORMULA) {
                    const formula = Formulas.processFormula(dataValidation.getCriteriaValues()[0].toString());
                    dataValidation = dataValidation.copy()
                        .requireFormulaSatisfied(formula)
                        .build();
                }
                range.setDataValidation(dataValidation);
            }
            info.conditionalFormats?.forEach(configurer => {
                if (configurer == null) {
                    return;
                }
                const originalConfigurer = configurer;
                configurer = builder => {
                    originalConfigurer(builder);
                    const formula = ConditionalFormatRuleUtils.extractFormula(builder);
                    if (formula != null) {
                        builder.whenFormulaSatisfied(Formulas.processFormula(formula));
                    }
                    return builder;
                };
                const fullRule = {
                    scope: conditionalFormattingScope,
                    order: ++conditionalFormattingOrder,
                    configurer,
                };
                ConditionalFormatting.addConditionalFormatRule(range, fullRule);
            });
        }
        sheet.getRange('1:1')
            .setHorizontalAlignment('center')
            .setFontWeight('bold')
            .setFontLine('none')
            .setNumberFormat('');
        DocumentFlags.set(this._documentFlag);
        DocumentFlags.cleanupByPrefix(this._documentFlagPrefix);
        const waitForAllDataExecutionsCompletion = SpreadsheetApp.getActiveSpreadsheet()['waitForAllDataExecutionsCompletion'];
        if (Utils.isFunction(waitForAllDataExecutionsCompletion)) {
            try {
                waitForAllDataExecutionsCompletion(5);
            }
            catch (e) {
                console.warn(e);
            }
        }
    }
}
class SheetLayoutProjects extends SheetLayout {
    get sheetName() {
        return GSheetProjectSettings.sheetName;
    }
    get columns() {
        return [
            {
                name: GSheetProjectSettings.iconColumnName,
                defaultTitleFontSize: 1,
                defaultWidth: '#default-height',
                defaultHorizontalAlignment: 'center',
            },
            
            {
                name: GSheetProjectSettings.milestoneColumnName,
                rangeName: GSheetProjectSettings.milestonesRangeName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.typeColumnName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.issueKeyColumnName,
                rangeName: GSheetProjectSettings.issuesRangeName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.childIssueKeyColumnName,
                rangeName: GSheetProjectSettings.childIssuesRangeName,
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(`=
                        #SELF_COLUMN(${GSheetProjectSettings.issuesRangeName})
                        =
                        OFFSET(#SELF_COLUMN(${GSheetProjectSettings.issuesRangeName}), -1, 0)
                    `)
                    .setHelpText(`Children should be grouped under their parent`)
                    .build(),
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.titleColumnName,
                rangeName: GSheetProjectSettings.titlesRangeName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.lastDataReloadColumnName,
                hiddenByDefault: true,
                defaultFormat: `yyyy-MM-dd HH:mm:ss.SSS`,
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.teamColumnName,
                rangeName: GSheetProjectSettings.teamsRangeName,
                //dataValidation <-- should be from ${GSheetProjectSettings.settingsTeamsTableTeamRangeName} range, see https://issuetracker.google.com/issues/143913035
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.estimateColumnName,
                rangeName: GSheetProjectSettings.estimatesRangeName,
                defaultFormat: '#,##0',
                defaultHorizontalAlignment: 'center',
                conditionalFormats: [
                    builder => builder
                        .whenFormulaSatisfied(`=
                            OR(
                                NOT(ISNUMBER(#SELF)),
                                #SELF < 0
                            )
                        `)
                        .setFontColor(GSheetProjectSettings.unimportantColor),
                    builder => builder
                        .whenFormulaSatisfied(`=
                            AND(
                                #SELF = "",
                                #SELF_COLUMN(${GSheetProjectSettings.teamsRangeName}) <> ""
                            )
                        `)
                        .setBackground(GSheetProjectSettings.importantWarningColor),
                ],
            },
            {
                name: GSheetProjectSettings.startColumnName,
                rangeName: GSheetProjectSettings.startsRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
                conditionalFormats: [
                    GSheetProjectSettings.inProgressesRangeName?.length
                        ? builder => builder
                            .whenFormulaSatisfied(`=
                                AND(
                                    #SELF <> "",
                                    #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}) <> "",
                                    #SELF > #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}),
                                    #SELF_COLUMN(${GSheetProjectSettings.inProgressesRangeName}) <> "",
                                    ISFORMULA(#SELF),
                                    FORMULATEXT(#SELF) <> "=TODAY()",
                                )
                            `)
                            .setBold(true)
                            .setFontColor(GSheetProjectSettings.errorColor)
                            .setItalic(true)
                            .setBackground(GSheetProjectSettings.unimportantWarningColor)
                        : null,
                    builder => builder
                        .whenFormulaSatisfied(`=
                            AND(
                                #SELF <> "",
                                #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}) <> "",
                                #SELF > #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}),
                                ISFORMULA(#SELF),
                                FORMULATEXT(#SELF) <> "=TODAY()",
                            )
                        `)
                        .setBold(true)
                        .setFontColor(GSheetProjectSettings.errorColor)
                        .setItalic(true),
                    GSheetProjectSettings.inProgressesRangeName?.length
                        ? builder => builder
                            .whenFormulaSatisfied(`=
                                AND(
                                    #SELF_COLUMN(${GSheetProjectSettings.inProgressesRangeName}) <> "",
                                    ISFORMULA(#SELF),
                                    FORMULATEXT(#SELF) <> "=TODAY()",
                                    #SELF <> ""
                                )
                            `)
                            .setItalic(true)
                            .setBackground(GSheetProjectSettings.unimportantWarningColor)
                        : null,
                ],
            },
            {
                name: GSheetProjectSettings.endColumnName,
                rangeName: GSheetProjectSettings.endsRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
                conditionalFormats: [
                    builder => builder
                        .whenFormulaSatisfied(`=
                            AND(
                                #SELF <> "",
                                #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName}) <> "",
                                #SELF > #SELF_COLUMN(${GSheetProjectSettings.deadlinesRangeName})
                            )
                        `)
                        .setBold(true)
                        .setFontColor(GSheetProjectSettings.errorColor),
                    GSheetProjectSettings.codeCompletesRangeName?.length
                        ? builder => builder
                            .whenFormulaSatisfied(`=
                                AND(
                                    #SELF_COLUMN(${GSheetProjectSettings.codeCompletesRangeName}) = "",
                                    #SELF <> "",
                                    #SELF < TODAY()
                                )
                            `)
                            .setBold(true)
                            .setFontColor(GSheetProjectSettings.warningColor)
                        : null,
                ],
            },
            {
                name: GSheetProjectSettings.earliestStartColumnName,
                rangeName: GSheetProjectSettings.earliestStartsRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.deadlineColumnName,
                rangeName: GSheetProjectSettings.deadlinesRangeName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            
        ];
    }
}
SheetLayoutProjects.instance = new SheetLayoutProjects();
class SheetLayouts {
    static get instances() {
        return [
            SheetLayoutProjects.instance,
            SheetLayoutSettings.instance,
        ];
    }
    static migrateIfNeeded() {
        Observability.timed([SheetLayouts.name, this.migrateIfNeeded.name].join(': '), () => {
            this.instances.forEach(instance => {
                const isMigrated = instance.migrateIfNeeded();
                if (isMigrated) {
                    this._isMigrated = true;
                }
            });
            this.applyAfterMigrationSteps();
        });
    }
    static migrate() {
        if (this._isMigrated) {
            return;
        }
        Observability.timed([SheetLayouts.name, this.migrate.name].join(': '), () => {
            this.instances.forEach(instance => instance.migrate());
            this.applyAfterMigrationSteps();
        });
    }
    static applyAfterMigrationSteps() {
        const rangeNames = [
            GSheetProjectSettings.issuesRangeName,
            GSheetProjectSettings.childIssuesRangeName,
            GSheetProjectSettings.titlesRangeName,
            GSheetProjectSettings.teamsRangeName,
            GSheetProjectSettings.estimatesRangeName,
            GSheetProjectSettings.startsRangeName,
            GSheetProjectSettings.endsRangeName,
            GSheetProjectSettings.deadlinesRangeName,
            GSheetProjectSettings.inProgressesRangeName,
            GSheetProjectSettings.codeCompletesRangeName,
            GSheetProjectSettings.settingsScheduleStartRangeName,
            GSheetProjectSettings.settingsScheduleBufferRangeName,
            GSheetProjectSettings.settingsTeamsTableRangeName,
            GSheetProjectSettings.settingsTeamsTableTeamRangeName,
            GSheetProjectSettings.settingsTeamsTableResourcesRangeName,
            GSheetProjectSettings.settingsMilestonesTableRangeName,
            GSheetProjectSettings.settingsMilestonesTableMilestoneRangeName,
            GSheetProjectSettings.settingsMilestonesTableDeadlineRangeName,
            GSheetProjectSettings.publicHolidaysRangeName,
        ].filter(it => it?.length).map(it => it);
        const missingRangeNames = rangeNames.filter(name => NamedRangeUtils.findNamedRange(name) == null);
        if (missingRangeNames.length) {
            throw new Error(`Missing named range(s): '${missingRangeNames.join("', '")}'`);
        }
        CommonFormatter.applyCommonFormatsToAllSheets();
    }
}
SheetLayouts._isMigrated = false;
class SheetLayoutSettings extends SheetLayout {
    get sheetName() {
        return GSheetProjectSettings.settingsSheetName;
    }
    get columns() {
        return [];
    }
}
SheetLayoutSettings.instance = new SheetLayoutSettings();
class SheetUtils {
    static findSheetByName(sheetName) {
        if (!sheetName?.length) {
            return undefined;
        }
        const sheets = ExecutionCache.getOrCompute('sheets-by-name', () => {
            const result = new Map();
            for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
                const name = Utils.normalizeName(sheet.getSheetName());
                result.set(name, sheet);
            }
            return result;
        }, `${SheetUtils.name}: ${this.findSheetByName.name}`);
        sheetName = Utils.normalizeName(sheetName);
        return sheets.get(sheetName);
    }
    static getSheetByName(sheetName) {
        return this.findSheetByName(sheetName) ?? (() => {
            throw new Error(`"${sheetName}" sheet can't be found`);
        })();
    }
    static isGridSheet(sheet) {
        if (Utils.isString(sheet)) {
            sheet = this.findSheetByName(sheet);
        }
        if (sheet == null) {
            return false;
        }
        return sheet.getType() === SpreadsheetApp.SheetType.GRID;
    }
    static getLastRow(sheet) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        return ExecutionCache.getOrCompute(['last-row', sheet], () => Math.max(sheet.getLastRow(), 1));
    }
    static setLastRow(sheet, lastRow) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        if (lastRow < 1) {
            lastRow = 1;
        }
        ExecutionCache.put(['last-row', sheet], lastRow);
    }
    static getLastColumn(sheet) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        return ExecutionCache.getOrCompute(['last-column', sheet], () => Math.max(sheet.getLastColumn(), 1));
    }
    static setLastColumn(sheet, lastColumn) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        if (lastColumn < 1) {
            lastColumn = 1;
        }
        ExecutionCache.put(['last-column', sheet], lastColumn);
    }
    static getMaxRows(sheet) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        return ExecutionCache.getOrCompute(['max-rows', sheet], () => Math.max(sheet.getMaxRows(), 1));
    }
    static setMaxRows(sheet, maxRows) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        if (maxRows < 1) {
            maxRows = 1;
        }
        const currentMaxRows = sheet.getMaxRows();
        if (currentMaxRows === maxRows) {
            // do nothing
        }
        else if (currentMaxRows < maxRows) {
            const rowsToInsert = maxRows - currentMaxRows;
            sheet.insertRowsAfter(currentMaxRows, rowsToInsert);
            sheet.getNamedRanges().forEach(namedRange => {
                const range = namedRange.getRange();
                if (range.getLastRow() >= currentMaxRows) {
                    const newRange = range.offset(0, 0, range.getNumRows() + rowsToInsert, range.getNumColumns());
                    namedRange.setRange(newRange);
                }
            });
            const filter = sheet.getFilter();
            if (filter != null) {
                const range = filter.getRange();
                if (range.getLastRow() >= currentMaxRows) {
                    filter.remove();
                    const newRange = range.offset(0, 0, range.getNumRows() + rowsToInsert, range.getNumColumns());
                    newRange.createFilter();
                }
            }
            const newConditionalFormatRules = sheet.getConditionalFormatRules().map(rule => {
                const newRanges = rule.getRanges().map(range => {
                    if (range.getLastRow() >= currentMaxRows) {
                        return range.offset(0, 0, range.getNumRows() + rowsToInsert, range.getNumColumns());
                    }
                    else {
                        return range;
                    }
                });
                return rule.copy().setRanges(newRanges);
            });
            sheet.setConditionalFormatRules(newConditionalFormatRules);
        }
        else {
            return; // do not reduce max rows
        }
        ExecutionCache.put(['max-rows', sheet], maxRows);
    }
    static getMaxColumns(sheet) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        return ExecutionCache.getOrCompute(['max-columns', sheet], () => Math.max(sheet.getMaxColumns(), 1));
    }
    static setMaxColumns(sheet, maxColumns) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        if (maxColumns < 1) {
            maxColumns = 1;
        }
        const currentMaxColumns = sheet.getMaxColumns();
        if (currentMaxColumns === maxColumns) {
            // do nothing
        }
        else if (currentMaxColumns < maxColumns) {
            const columnsToInsert = maxColumns - currentMaxColumns;
            sheet.insertColumnsAfter(currentMaxColumns, columnsToInsert);
        }
        else {
            return; // do not reduce max columns
        }
        ExecutionCache.put(['max-columns', sheet], maxColumns);
    }
    static getWholeSheetRange(sheet) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        return sheet.getRange(1, 1, SheetUtils.getMaxRows(sheet), SheetUtils.getMaxColumns(sheet));
    }
    static findColumnByName(sheet, columnName) {
        if (!columnName?.length) {
            return undefined;
        }
        if (Utils.isString(sheet)) {
            sheet = this.findSheetByName(sheet);
        }
        if (sheet == null || !this.isGridSheet(sheet)) {
            return undefined;
        }
        ProtectionLocks.lockAllColumns(sheet);
        const columns = ExecutionCache.getOrCompute(['columns-by-name', sheet], () => {
            const result = new Map();
            for (const col of Utils.range(GSheetProjectSettings.titleRow, this.getLastColumn(sheet))) {
                const name = Utils.normalizeName(sheet.getRange(1, col).getValue());
                result.set(name, col);
            }
            return result;
        }, `${SheetUtils.name}: ${this.findColumnByName.name}`);
        columnName = Utils.normalizeName(columnName);
        return columns.get(columnName);
    }
    static getColumnByName(sheet, columnName) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        return this.findColumnByName(sheet, columnName) ?? (() => {
            throw new Error(`"${sheet.getSheetName()}" sheet: "${columnName}" column can't be found`);
        })();
    }
    static getColumnRange(sheet, column, minRow) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        if (Utils.isString(column)) {
            column = this.getColumnByName(sheet, column);
        }
        if (minRow == null || minRow < 1) {
            minRow = 1;
        }
        const lastRow = this.getLastRow(sheet);
        if (minRow > lastRow) {
            return sheet.getRange(minRow, column);
        }
        const rows = lastRow - minRow + 1;
        return sheet.getRange(minRow, column, rows, 1);
    }
    static getColumnsValues(sheet, columns, minRow, maxRow) {
        function getValues(range) {
            return range.getValues();
        }
        return this._getColumnsProps(sheet, columns, getValues, minRow, maxRow);
    }
    static getColumnsStringValues(sheet, columns, minRow, maxRow) {
        function getValues(range) {
            return range.getValues();
        }
        const result = this._getColumnsProps(sheet, columns, getValues, minRow, maxRow);
        for (const [key, values] of Object.entries(result)) {
            result[key] = values.map(value => value.toString());
        }
        return result;
    }
    static getColumnsFormulas(sheet, columns, minRow, maxRow) {
        function getFormulas(range) {
            return range.getFormulas();
        }
        return this._getColumnsProps(sheet, columns, getFormulas, minRow, maxRow);
    }
    static _getColumnsProps(sheet, columns, getter, minRow, maxRow) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        if (!Object.keys(columns).length) {
            return {};
        }
        if (minRow == null || minRow < 1) {
            minRow = 1;
        }
        if (maxRow == null) {
            maxRow = this.getLastRow(sheet);
        }
        const columnToNumber = Object.keys(columns)
            .reduce((rec, key) => {
            const value = columns[key];
            rec[key] = Utils.isString(value)
                ? this.getColumnByName(sheet, value)
                : value;
            return rec;
        }, {});
        const numbers = Object.values(columnToNumber).filter(Utils.distinct()).toSorted(Utils.numericAsc());
        const result = {};
        Object.keys(columns).forEach(key => result[key] = []);
        if (minRow > maxRow) {
            return result;
        }
        Observability.timed([
            SheetUtils.name,
            this._getColumnsProps.name,
            sheet.getSheetName(),
            `rows from #${minRow} to #${maxRow}`,
            `columns #${numbers.join(', #')} (${getter.name})`,
        ].join(': '), () => {
            while (numbers.length) {
                const baseColumn = numbers.shift();
                let columnsCount = 1;
                while (numbers.length) {
                    const nextNumber = numbers[0];
                    if (nextNumber === baseColumn + columnsCount) {
                        ++columnsCount;
                        numbers.shift();
                    }
                    else {
                        break;
                    }
                }
                const range = sheet.getRange(minRow, baseColumn, Math.max(maxRow - minRow + 1, 1), columnsCount);
                const props = getter(range);
                props.forEach(rows => rows.forEach((columnValue, index) => {
                    const column = baseColumn + index;
                    for (const [columnKey, columnNumber] of Object.entries(columnToNumber)) {
                        if (column === columnNumber) {
                            result[columnKey].push(columnValue);
                        }
                    }
                }));
            }
        });
        return result;
    }
    static getRowRange(sheet, row, minColumn) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        if (minColumn == null) {
            minColumn = 1;
        }
        else if (Utils.isString(minColumn)) {
            minColumn = this.getColumnByName(sheet, minColumn);
        }
        else if (minColumn < 1) {
            minColumn = 1;
        }
        const lastColumn = this.getLastColumn(sheet);
        if (minColumn > lastColumn) {
            return sheet.getRange(row, minColumn);
        }
        const columns = lastColumn - minColumn + 1;
        return sheet.getRange(row, minColumn, 1, columns);
    }
}

class Utils {
    static *range(startIncluding, endIncluding) {
        for (let n = startIncluding; n <= endIncluding; ++n) {
            yield n;
        }
    }
    static normalizeName(name) {
        return name.toString().trim().replaceAll(/\s+/g, ' ').toLowerCase();
    }
    static toLowerCamelCase(value) {
        value = value.replace(/^[^a-z0-9]+/i, '').replace(/[^a-z0-9]+$/i, '');
        if (value.length <= 1) {
            return value.toLowerCase();
        }
        value = value.substring(0, 1).toLowerCase() + value.substring(1).toLowerCase();
        value = value.replaceAll(/[^a-z0-9]+([a-z0-9])/ig, (_, letter) => letter.toUpperCase());
        return value;
    }
    static toUpperCamelCase(value) {
        value = this.toLowerCamelCase(value);
        if (value.length <= 1) {
            return value.toUpperCase();
        }
        value = value.substring(0, 1).toUpperCase() + value.substring(1);
        return value;
    }
    static mapRecordValues(record, transformer) {
        const result = {};
        Object.entries(record).forEach(([key, value]) => {
            result[key] = transformer(value, key);
        });
        return result;
    }
    static mapToRecord(keys, transformer) {
        const result = {};
        keys.forEach(key => {
            result[key] = transformer(key);
        });
        return result;
    }
    /**
     * See https://stackoverflow.com/a/44134328/3740528
     */
    static hslToRgb(h, s, l) {
        l /= 100;
        const a = s * Math.min(l, 1 - l) / 100;
        const f = (n) => {
            const k = (n + h / 30) % 12;
            const color = l - a * Math.max(Math.min(k - 3, 9 - k, 1), -1);
            return Math.round(255 * color).toString(16).padStart(2, '0');
        };
        return `#${f(0)}${f(8)}${f(4)}`;
    }
    static toJsonObject(object, callGetters = true, keepNulls = false) {
        if (object == null) {
            return object;
        }
        else if (object instanceof Date) {
            return new Date(object.getTime());
        }
        else if (Array.isArray(object)) {
            const result = [];
            for (const element of object) {
                if (element === object || element === undefined || (!keepNulls && element === null)) {
                    continue;
                }
                result.push(this.toJsonObject(element, callGetters));
            }
            return result;
        }
        else if (this.isFunction(object.toJSON)) {
            return object.toJSON();
        }
        else if (this.isFunction(object.getA1Notation)) {
            return object.getA1Notation();
        }
        else if (this.isFunction(object.getSheetName)) {
            return object.getSheetName();
        }
        else if (typeof object === 'object') {
            const prototypePropertiesToExclude = ['constructor'];
            const properties = [];
            for (const property in object) {
                if (!object.hasOwnProperty(property)
                    && prototypePropertiesToExclude.includes(property)) {
                    continue;
                }
                properties.push(property);
            }
            properties.sort((p1, p2) => {
                const n1 = parseFloat(p1);
                const n2 = parseFloat(p2);
                if (!isNaN(n1) && !isNaN(n2)) {
                    return n1 - n2;
                }
                return p1.localeCompare(p2);
            });
            const result = {};
            for (const property of properties) {
                const value = object[property];
                if (value === object || value === undefined || (!keepNulls && value === null)) {
                    continue;
                }
                if (this.isFunction(value)) {
                    if (callGetters) {
                        const getterMatcher = property.match(/^(get|is)([A-Z].*)$/);
                        if (getterMatcher != null) {
                            const propValue = value.call(object);
                            if (propValue === object || propValue === undefined || (!keepNulls && propValue === null)) {
                                continue;
                            }
                            let name = getterMatcher[2];
                            name = name.substring(0, 1).toLowerCase() + name.substring(1);
                            result[name] = this.toJsonObject(propValue, callGetters);
                        }
                    }
                    continue;
                }
                result[property] = this.toJsonObject(value, callGetters);
            }
            return result;
        }
        else {
            return object;
        }
    }
    static groupBy(array, keyGetter) {
        const result = new Map();
        for (const element of array) {
            const key = keyGetter(element);
            if (key != null) {
                let groupedElements = result.get(key);
                if (groupedElements == null) {
                    groupedElements = [];
                    result.set(key, groupedElements);
                }
                groupedElements.push(element);
            }
        }
        return result;
    }
    static extractRegex(string, regexp, group) {
        if (string == null) {
            return null;
        }
        if (this.isString(regexp)) {
            regexp = new RegExp(regexp);
        }
        if (group == null) {
            group = 0;
        }
        const match = regexp.exec(string);
        if (match == null) {
            return null;
        }
        if (match.groups != null) {
            const result = match.groups[group];
            if (result != null) {
                return result;
            }
        }
        return match[group];
    }
    static distinct() {
        return (value, index, array) => array.indexOf(value) === index;
    }
    static distinctBy(getter) {
        const seen = new Set();
        return (value) => {
            const property = getter(value);
            if (seen.has(property)) {
                return false;
            }
            seen.add(property);
            return true;
        };
    }
    static trimArrayEndBy(array, predicate) {
        while (array.length) {
            const lastElement = array[array.length - 1];
            if (predicate(lastElement)) {
                --array.length;
            }
            else {
                break;
            }
        }
    }
    static arrayRemoveIf(array, predicate) {
        for (let index = 0; index < array.length; ++index) {
            const element = array[index];
            if (predicate(element)) {
                array.splice(index, 1);
                --index;
            }
        }
    }
    static moveArrayElements(array, fromIndex, count, targetIndex) {
        if (fromIndex === targetIndex || count <= 0) {
            return;
        }
        if (array.length <= targetIndex) {
            array.length = targetIndex + 1;
        }
        const moved = array.splice(fromIndex, count);
        array.splice(targetIndex, 0, ...moved);
    }
    static parseDate(value) {
        if (value == null) {
            return null;
        }
        else if (this.isNumber(value)) {
            return new Date(value);
        }
        else if (Utils.isString(value)) {
            try {
                return new Date(Number.isNaN(value) ? value : parseFloat(value));
            }
            catch (_) {
                return null;
            }
        }
        else if (this.isFunction(value.getTime)) {
            return this.parseDate(value.getTime());
        }
        else {
            return null;
        }
    }
    static parseDateOrThrow(value) {
        return this.parseDate(value) ?? (() => {
            throw new Error(`Not a date: "${value}"`);
        })();
    }
    static hashCode(value) {
        if (!value?.length) {
            return 0;
        }
        let hash = 0;
        for (let i = 0; i < value.length; ++i) {
            const chr = value.charCodeAt(i);
            hash = ((hash << 5) - hash) + chr;
            hash |= 0;
        }
        return hash;
    }
    static merge(...objects) {
        const result = {};
        for (const object of objects) {
            if (object == null) {
                continue;
            }
            for (const key of Object.keys(object)) {
                const value = object[key];
                if (value === undefined) {
                    continue;
                }
                const currentValue = result[key];
                if (this.isRecord(value) && this.isRecord(currentValue)) {
                    result[key] = this.merge(currentValue, value);
                    continue;
                }
                result[key] = value;
            }
        }
        return result;
    }
    static mergeInto(result, ...objects) {
        for (const object of objects) {
            if (object == null) {
                continue;
            }
            for (const key of Object.keys(object)) {
                const value = object[key];
                if (value === undefined) {
                    continue;
                }
                const currentValue = result[key];
                if (this.isRecord(value) && this.isRecord(currentValue)) {
                    this.mergeInto(currentValue, value);
                    continue;
                }
                result[key] = value;
            }
        }
        return result;
    }
    static arrayEquals(array1, array2) {
        if (array1 === array2) {
            return true;
        }
        else if (array1 == null || array2 == null) {
            return false;
        }
        if (array1.length !== array2.length) {
            return false;
        }
        for (let i = 0; i < array1.length; ++i) {
            const element1 = array1[i];
            const element2 = array2[i];
            if (element1 !== element2) {
                return false;
            }
        }
        return true;
    }
    static arrayOf(length, initValue) {
        const array = new Array(length ?? 0);
        if (initValue !== undefined) {
            array.fill(initValue);
        }
        return array;
    }
    static escapeRegex(string) {
        return string.replaceAll(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    static numericAsc() {
        return (n1, n2) => n1 - n2;
    }
    static numericDesc() {
        return (n1, n2) => n2 - n1;
    }
    static isString(value) {
        return typeof value === 'string';
    }
    static isNumber(value) {
        return typeof value === 'number';
    }
    static isBoolean(value) {
        return typeof value === 'boolean';
    }
    static isFunction(value) {
        return typeof value === 'function';
    }
    static isRecord(value) {
        return typeof value === 'object' && !Array.isArray(value);
    }
    static throwNotImplemented(...name) {
        throw new Error(`Not implemented: ${name.join(': ')}`);
    }
}