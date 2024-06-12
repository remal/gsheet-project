/**
 * SHA-256 digest of the provided input
 * @param {unknown} value
 * @returns {string}
 * @customFunction
 */
function SHA256(value) {
    var _a;
    const string = (_a = value === null || value === void 0 ? void 0 : value.toString()) !== null && _a !== void 0 ? _a : '';
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
    });
}
function applyDefaultStylesOfGSheetProject() {
    EntryPoint.entryPoint(() => {
        SheetLayouts.migrate();
    });
}
function onOpenGSheetProject(event) {
    SpreadsheetApp.getUi()
        .createMenu("GSheetProject")
        .addItem("Refresh selected rows", refreshSelectedRowsOfGSheetProject.name)
        .addItem("Refresh all rows", refreshAllRowsOfGSheetProject.name)
        .addItem("Apply default styles", applyDefaultStylesOfGSheetProject.name)
        .addToUi();
    EntryPoint.entryPoint(() => {
        SheetLayouts.migrateIfNeeded();
    });
}
function onChangeGSheetProject(event) {
    var _a, _b;
    function onInsert() {
        EntryPoint.entryPoint(() => {
            CommonFormatter.applyCommonFormatsToAllSheets();
        });
    }
    function onRemove() {
        applyDefaultStylesOfGSheetProject();
    }
    const changeType = (_b = (_a = event === null || event === void 0 ? void 0 : event.changeType) === null || _a === void 0 ? void 0 : _a.toString()) !== null && _b !== void 0 ? _b : '';
    if (['INSERT_ROW', 'INSERT_COLUMN'].includes(changeType)) {
        onInsert();
    }
    else if (['REMOVE_COLUMN'].includes(changeType)) {
        onRemove();
    }
}
function onEditGSheetProject(event) {
    const range = event === null || event === void 0 ? void 0 : event.range;
    if (range == null) {
        return;
    }
    EntryPoint.entryPoint(() => {
        //Utils.timed(`Done logic`, () => DoneLogic.executeDoneLogic(range))
        Utils.timed(`Issue hierarchy`, () => IssueHierarchyFormatter.formatHierarchy(range));
        Utils.timed(`Default formulas`, () => DefaultFormulas.insertDefaultFormulas(range));
        Utils.timed(`Reload issue data`, () => IssueDataDisplay.reloadIssueData(range));
    });
}
function onFormSubmitGSheetProject(event) {
    onEditGSheetProject({
        range: event === null || event === void 0 ? void 0 : event.range,
    });
}
var _a;
class GSheetProjectSettings {
    static computeStringSettingsHash() {
        const hashableValues = {};
        for (const [key, value] of Object.entries(_a)) {
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
GSheetProjectSettings.useLoadingImage = false;
GSheetProjectSettings.skipHiddenIssues = true;
//static restoreUndoneEnd: boolean = false
GSheetProjectSettings.issuesRangeName = 'Issues';
GSheetProjectSettings.childIssuesRangeName = 'ChildIssues';
GSheetProjectSettings.teamsRangeName = "Teams";
GSheetProjectSettings.settingsTeamsTableRangeName = 'TeamsTable';
GSheetProjectSettings.settingsTeamsTableTeamRangeName = 'TeamsTableTeam';
GSheetProjectSettings.settingsTeamsTableResourcesRangeName = 'TeamsTableResources';
GSheetProjectSettings.issueTrackers = [];
GSheetProjectSettings.issuesLoadTimeoutMillis = 5 * 60 * 1000;
GSheetProjectSettings.booleanIssuesMetrics = {};
GSheetProjectSettings.counterIssuesMetrics = {};
GSheetProjectSettings.sheetName = "Projects";
GSheetProjectSettings.iconColumnName = "icon";
//static doneColumnName: ColumnName = "Done"
GSheetProjectSettings.milestoneColumnName = "Milestone";
GSheetProjectSettings.typeColumnName = "Type";
GSheetProjectSettings.issueColumnName = "Issue";
GSheetProjectSettings.childIssueColumnName = "Child\nIssue";
GSheetProjectSettings.lastDataReloadColumnName = "Last\nReload";
GSheetProjectSettings.titleColumnName = "Title";
GSheetProjectSettings.teamColumnName = "Team";
GSheetProjectSettings.estimateColumnName = "Estimate\n(days)";
GSheetProjectSettings.deadlineColumnName = "Deadline";
GSheetProjectSettings.startColumnName = "Start";
GSheetProjectSettings.endColumnName = "End";
//static issueHashColumnName: ColumnName = "Issue Hash"
GSheetProjectSettings.settingsSheetName = "Settings";
GSheetProjectSettings.settingsScheduleStartRangeName = 'ScheduleStart';
GSheetProjectSettings.settingsScheduleBufferRangeName = 'ScheduleBuffer';
GSheetProjectSettings.indent = 4;
class AbstractIssueLogic {
    static _processRange(range) {
        if (![
            GSheetProjectSettings.issueColumnName,
            GSheetProjectSettings.childIssueColumnName,
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
            issues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueColumnName),
            childIssues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName),
        }, startRow, endRow);
        Utils.trimArrayEndBy(result.issues, it => !(it === null || it === void 0 ? void 0 : it.length));
        result.childIssues.length = result.issues.length;
        return result;
    }
    static _getIssueValuesWithLastReloadDate(range) {
        const sheet = range.getSheet();
        const startRow = range.getRow();
        const endRow = startRow + range.getNumRows() - 1;
        const result = SheetUtils.getColumnsValues(sheet, {
            issues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueColumnName),
            childIssues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName),
            lastDataReload: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.lastDataReloadColumnName),
        }, startRow, endRow);
        Utils.trimArrayEndBy(result.issues, it => { var _a; return !((_a = it === null || it === void 0 ? void 0 : it.toString()) === null || _a === void 0 ? void 0 : _a.length); });
        result.childIssues.length = result.issues.length;
        result.lastDataReload.length = result.issues.length;
        return {
            issues: result.issues.map(it => it === null || it === void 0 ? void 0 : it.toString()),
            childIssues: result.childIssues.map(it => it === null || it === void 0 ? void 0 : it.toString()),
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
            this.setMiddleVerticalAlign(sheet);
            this.setClipWrapStrategy(sheet);
            this.highlightCellsWithFormula(sheet);
        });
    }
    static setMiddleVerticalAlign(sheet) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet);
        }
        sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
            .setVerticalAlignment('middle');
    }
    static setClipWrapStrategy(sheet) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet);
        }
        sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }
    static highlightCellsWithFormula(sheet) {
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet);
        }
        ConditionalFormatting.addConditionalFormatRule(sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()), {
            scope: 'common',
            order: 10000,
            configurer: builder => builder
                .whenFormulaSatisfied('=ISFORMULA(A1)')
                .setItalic(true)
                .setFontColor('#333'),
        });
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
    static addConditionalFormatRule(range, orderedRule) {
        var _a;
        if (!GSheetProjectSettings.updateConditionalFormatRules) {
            return;
        }
        const builder = SpreadsheetApp.newConditionalFormatRule();
        builder.setRanges([range]);
        orderedRule.configurer(builder);
        let formula = ConditionalFormatRuleUtils.extractFormula(builder);
        if (formula == null) {
            throw new Error(`Not a boolean condition with formula`);
        }
        formula = '=AND(' + [
            Utils.processFormula(formula
                .replace(/^=/, '')
                .replace(/^and\(\s*(.+)\s*\)$/i, '$1')),
            `"GSPs"<>"${orderedRule.scope}"`,
            `"GSPo"<>"${orderedRule.order}"`,
        ].join(', ') + ')';
        builder.whenFormulaSatisfied(formula);
        const newRule = builder.build();
        const sheet = range.getSheet();
        let rules = (_a = sheet.getConditionalFormatRules()) !== null && _a !== void 0 ? _a : [];
        rules = rules.filter(rule => !(this._extractScope(rule) === orderedRule.scope && this._extractOrder(rule) === orderedRule.order));
        rules.push(newRule);
        rules = rules.toSorted((r1, r2) => {
            const o1 = this._extractOrder(r1);
            const o2 = this._extractOrder(r2);
            if (o1 === null && o2 === null) {
                return 0;
            }
            else if (o2 !== null) {
                return 1;
            }
            else if (o1 !== null) {
                return 11;
            }
            else {
                return o2 - o1;
            }
        });
        sheet.setConditionalFormatRules(rules);
    }
    static removeConditionalFormatRulesByScope(sheet, scopeToRemove) {
        var _a;
        if (Utils.isString(sheet)) {
            sheet = SheetUtils.getSheetByName(sheet);
        }
        const rules = (_a = sheet.getConditionalFormatRules()) !== null && _a !== void 0 ? _a : [];
        const filteredRules = rules.filter(rule => this._extractScope(rule) !== scopeToRemove);
        if (filteredRules.length !== rules.length) {
            sheet.setConditionalFormatRules(filteredRules);
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
        const match = rule.match(/^=(?:AND|and)\(.+, "GSPs"\s*<>\s*"([a-z]*)"/);
        if (match) {
            return match[1];
        }
        return undefined;
    }
    static _extractOrder(rule) {
        if (!Utils.isString(rule)) {
            const formula = ConditionalFormatRuleUtils.extractFormula(rule);
            if (formula == null) {
                return undefined;
            }
            rule = formula;
        }
        const match = rule.match(/^=(?:AND|and)\(.+, "GSPo"\s*<>\s*"(\d+(\.\d*)?)"/);
        if (match) {
            return parseFloat(match[1]);
        }
        return undefined;
    }
}
class DefaultFormulas extends AbstractIssueLogic {
    static insertDefaultFormulas(range) {
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
        const { issues, childIssues } = this._getIssueValues(range);
        const addFormulas = (column, formulaGenerator) => Utils.timed([
            DefaultFormulas.name,
            sheet.getSheetName(),
            addFormulas.name,
            `column #${column}`,
        ].join(': '), () => {
            var _a, _b, _c, _d, _e;
            const values = this._getStringValues(range, column);
            const formulas = this._getFormulas(range, column);
            for (let row = startRow; row <= endRow; ++row) {
                const index = row - startRow;
                if (!((_a = issues[index]) === null || _a === void 0 ? void 0 : _a.length) && !((_b = childIssues[index]) === null || _b === void 0 ? void 0 : _b.length)) {
                    if ((_c = formulas[index]) === null || _c === void 0 ? void 0 : _c.length) {
                        sheet.getRange(row, column).setFormula('');
                    }
                    continue;
                }
                if (!((_d = values[index]) === null || _d === void 0 ? void 0 : _d.length) && !((_e = formulas[index]) === null || _e === void 0 ? void 0 : _e.length)) {
                    console.info([
                        DefaultFormulas.name,
                        sheet.getSheetName(),
                        addFormulas.name,
                        `column #${column}`,
                        `row #${row}`,
                    ].join(': '));
                    const formula = Utils.processFormula(formulaGenerator(row));
                    sheet.getRange(row, column).setFormula(formula);
                }
            }
        });
        const teamColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.teamColumnName);
        const estimateColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName);
        const startColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.startColumnName);
        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName);
        addFormulas(startColumn, row => {
            const teamFirstA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(GSheetProjectSettings.firstDataRow, teamColumn));
            const estimateFirstA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(GSheetProjectSettings.firstDataRow, estimateColumn));
            const endFirstA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(GSheetProjectSettings.firstDataRow, endColumn));
            const teamA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, teamColumn));
            const estimateA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, estimateColumn));
            const resourcesLookup = `
                VLOOKUP(
                    ${teamA1Notation},
                    ${GSheetProjectSettings.settingsTeamsTableRangeName},
                    1
                        + COLUMN(${GSheetProjectSettings.settingsTeamsTableResourcesRangeName})
                        - COLUMN(${GSheetProjectSettings.settingsTeamsTableRangeName}),
                    FALSE
                )
            `;
            const notEnoughPreviousLanes = `
                COUNTIFS(
                    OFFSET(
                        ${teamFirstA1Notation},
                        0,
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    "=" & ${teamA1Notation},
                    OFFSET(
                        ${estimateFirstA1Notation},
                        0,
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    ">0"
                ) < ${resourcesLookup}
            `;
            const filter = `
                FILTER(
                    OFFSET(
                        ${endFirstA1Notation},
                        0,
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ),
                    OFFSET(
                        ${teamFirstA1Notation},
                        0,
                        0,
                        ROW() - ${GSheetProjectSettings.firstDataRow},
                        1
                    ) = ${teamA1Notation},
                    OFFSET(
                        ${estimateFirstA1Notation},
                        0,
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
                        ${resourcesLookup},
                        0,
                        1,
                        FALSE
                    )
                )
            `;
            const nextWorkdayLastEnd = `
                WORKDAY(${lastEnd}, 1)
            `;
            const firstDataRowIf = `
                IF(
                    OR(
                        ROW() <= ${GSheetProjectSettings.firstDataRow},
                        ${notEnoughPreviousLanes}
                    ),
                    ${GSheetProjectSettings.settingsScheduleStartRangeName},
                    ${nextWorkdayLastEnd}
                )
            `;
            const notEnoughDataIf = `
                IF(
                    OR(
                        ${teamA1Notation} = "",
                        ${estimateA1Notation} = ""
                    ),
                    "",
                    ${firstDataRowIf}
                )
            `;
            return `=${notEnoughDataIf}`;
        });
        addFormulas(endColumn, row => {
            const startA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, startColumn));
            const estimateA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, estimateColumn));
            const bufferRangeName = GSheetProjectSettings.settingsScheduleBufferRangeName;
            return `
                =IF(
                    OR(
                        ${startA1Notation} = "",
                        ${estimateA1Notation} = ""
                    ),
                    "",
                    WORKDAY(${startA1Notation}, ROUND(${estimateA1Notation} * (1 + ${bufferRangeName})))
                )
            `;
        });
    }
}
class DocumentFlags {
    static set(key, value = true) {
        if (value) {
            PropertiesService.getDocumentProperties().setProperty(key, new Date().getTime().toString());
        }
        else {
            PropertiesService.getDocumentProperties().deleteProperty(key);
        }
    }
    static isSet(key) {
        var _a;
        return (_a = PropertiesService.getDocumentProperties().getProperty(key)) === null || _a === void 0 ? void 0 : _a.length;
    }
    static cleanupByPrefix(keyPrefix) {
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
    static entryPoint(action) {
        if (this._isInEntryPoint) {
            return action();
        }
        try {
            this._isInEntryPoint = true;
            ExecutionCache.resetCache();
            return action();
        }
        catch (e) {
            console.error(e);
            SpreadsheetApp.getActiveSpreadsheet().toast(e.toString(), "Automation error");
            throw e;
        }
        finally {
            ProtectionLocks.release();
            ProtectionLocks.releaseExpiredLocks();
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
        if (timerLabel === null || timerLabel === void 0 ? void 0 : timerLabel.length) {
            console.time(timerLabel);
        }
        let result;
        try {
            result = compute();
        }
        finally {
            if (timerLabel === null || timerLabel === void 0 ? void 0 : timerLabel.length) {
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
            if (Utils.isFunction(value.getUniqueId)) {
                return value.getUniqueId();
            }
            else if (Utils.isFunction(value.getSheetId)) {
                return value.getSheetId();
            }
            else if (Utils.isFunction(value.getId)) {
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
class Images {
}
Images.loadingImageUrl = 'https://raw.githubusercontent.com/remal/misc/main/spinner-100.gif';
class IssueDataDisplay extends AbstractIssueLogic {
    static reloadIssueData(range) {
        var _a, _b;
        const processedRange = this._processRange(range);
        if (processedRange == null) {
            return;
        }
        else {
            range = processedRange;
        }
        const sheet = range.getSheet();
        const iconColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.iconColumnName);
        const issueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueColumnName);
        const childIssueColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName);
        const lastDataReloadColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.lastDataReloadColumnName);
        const titleColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.titleColumnName);
        const { issues, childIssues, lastDataReload } = this._getIssueValuesWithLastReloadDate(range);
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
        const start = Date.now();
        for (const index of indexes) {
            if (Date.now() - start >= GSheetProjectSettings.issuesLoadTimeoutMillis) {
                const message = "Issues load timeout occurred";
                console.warn(message);
                //SpreadsheetApp.getActiveSpreadsheet().toast(message)
                break;
            }
            const row = range.getRow() + index;
            const cleanupColumns = () => {
                const notations = [
                    [
                        sheet.getRange(row, titleColumn),
                        sheet.getRange(row, iconColumn),
                    ],
                    Object.keys(GSheetProjectSettings.booleanIssuesMetrics)
                        .map(columnName => SheetUtils.findColumnByName(sheet, columnName))
                        .filter(column => column != null)
                        .map(column => sheet.getRange(row, column)),
                    Object.keys(GSheetProjectSettings.counterIssuesMetrics)
                        .map(columnName => SheetUtils.findColumnByName(sheet, columnName))
                        .filter(column => column != null)
                        .map(column => sheet.getRange(row, column)),
                ]
                    .flat()
                    .map(range => range.getA1Notation());
                if (notations.length) {
                    sheet.getRangeList(notations).setValue('');
                }
                sheet.getRange(row, lastDataReloadColumn).setValue(new Date());
            };
            if (GSheetProjectSettings.skipHiddenIssues && sheet.isRowHiddenByUser(row)) { // a slow check
                cleanupColumns();
                continue;
            }
            if (GSheetProjectSettings.useLoadingImage) {
                sheet.getRange(row, iconColumn).setFormula(`=IMAGE("${Images.loadingImageUrl}")`);
            }
            else {
                sheet.getRange(row, iconColumn).setValue('...');
            }
            let currentIssueColumn;
            let originalIssueKeysText;
            if ((_a = childIssues[index]) === null || _a === void 0 ? void 0 : _a.length) {
                currentIssueColumn = childIssueColumn;
                originalIssueKeysText = childIssues[index];
            }
            else if ((_b = issues[index]) === null || _b === void 0 ? void 0 : _b.length) {
                currentIssueColumn = issueColumn;
                originalIssueKeysText = issues[index];
            }
            else {
                cleanupColumns();
                continue;
            }
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
                if (issueId === null || issueId === void 0 ? void 0 : issueId.length) {
                    issueKeyIds[issueKey] = issueId;
                }
                const searchQuery = issueTracker.extractSearchQuery(issueKey);
                if (searchQuery === null || searchQuery === void 0 ? void 0 : searchQuery.length) {
                    issueKeyQueries[issueKey] = searchQuery;
                }
            }
            if (issueTracker == null) {
                cleanupColumns();
                continue;
            }
            const allIssueLinks = allIssueKeys.map(issueKey => {
                if (issueKeys.includes(issueKey)) {
                    const issueId = issueKeyIds[issueKey];
                    if (issueId === null || issueId === void 0 ? void 0 : issueId.length) {
                        return {
                            title: issueTracker.canonizeIssueKey(issueKey),
                            url: issueTracker.getUrlForIssueId(issueId),
                        };
                    }
                    else {
                        const searchQuery = issueKeyQueries[issueKey];
                        if (searchQuery === null || searchQuery === void 0 ? void 0 : searchQuery.length) {
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
            sheet.getRange(row, currentIssueColumn).setRichTextValue(RichTextUtils.createLinksValue(allIssueLinks));
            const loadedIssues = LazyProxy.create(() => Utils.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading issues`,
            ].join(': '), () => {
                const issueIds = Object.values(issueKeyIds).filter(Utils.distinct());
                return issueTracker.loadIssues(issueIds);
            }));
            const loadedChildIssues = LazyProxy.create(() => Utils.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading child issues`,
            ].join(': '), () => {
                const issueIds = loadedIssues.map(it => it.id);
                return [
                    issueTracker.loadChildren(issueIds),
                    Object.values(issueKeyQueries)
                        .filter(Utils.distinct())
                        .flatMap(query => issueTracker.search(query)),
                ]
                    .flat()
                    .filter(Utils.distinctBy(issue => issue.id))
                    .filter(issue => !issueIds.includes(issue.id));
            }));
            const loadedBlockerIssues = LazyProxy.create(() => Utils.timed([
                IssueDataDisplay.name,
                this.reloadIssueData.name,
                `row #${row}`,
                `loading blocker issues`,
            ].join(': '), () => {
                const allIssueIds = [loadedIssues, loadedChildIssues]
                    .flatMap(it => it.map(it => it.id))
                    .filter(Utils.distinct());
                return issueTracker.loadBlockers(allIssueIds)
                    .filter(issue => !allIssueIds.includes(issue.id));
            }));
            const titles = issueKeys.map(issueKey => {
                var _a, _b;
                const issueId = issueKeyIds[issueKey];
                if (issueId === null || issueId === void 0 ? void 0 : issueId.length) {
                    return (_a = loadedIssues.find(issue => issue.id)) === null || _a === void 0 ? void 0 : _a.title;
                }
                if ((_b = issueKeyQueries[issueKey]) === null || _b === void 0 ? void 0 : _b.length) {
                    return Utils.timed([
                        IssueDataDisplay.name,
                        this.reloadIssueData.name,
                        `row #${row}`,
                        `loading search title for "${issueKey}" issue key`,
                    ].join(': '), () => issueTracker.loadIssueKeySearchTitle(issueKey));
                }
                return undefined;
            })
                .map(title => title === null || title === void 0 ? void 0 : title.trim())
                .filter(title => title === null || title === void 0 ? void 0 : title.length)
                .map(title => title);
            sheet.getRange(row, titleColumn).setValue(titles.join('\n'));
            for (const [columnName, issuesMetric] of Object.entries(GSheetProjectSettings.booleanIssuesMetrics)) {
                const column = SheetUtils.findColumnByName(sheet, columnName);
                if (column == null) {
                    continue;
                }
                const value = issuesMetric(loadedIssues, loadedChildIssues, loadedBlockerIssues);
                sheet.getRange(row, column).setValue(value ? "Yes" : '');
            }
            for (const [columnName, issuesCounterMetric] of Object.entries(GSheetProjectSettings.counterIssuesMetrics)) {
                const column = SheetUtils.findColumnByName(sheet, columnName);
                if (column == null) {
                    continue;
                }
                const foundIssues = issuesCounterMetric(loadedIssues, loadedChildIssues, loadedBlockerIssues);
                if (!foundIssues.length) {
                    sheet.getRange(row, column).setValue('');
                    continue;
                }
                const foundIssueIds = foundIssues.map(it => it.id)
                    .filter(Utils.distinct());
                const link = {
                    title: foundIssues.length.toString(),
                    url: issueTracker.getUrlForIssueIds(foundIssueIds),
                };
                sheet.getRange(row, column).setRichTextValue(RichTextUtils.createLinkValue(link));
            }
            sheet.getRange(row, lastDataReloadColumn).setValue(allIssueKeys.length ? new Date() : '');
            sheet.getRange(row, iconColumn).setValue('');
        }
    }
    static reloadAllIssuesData() {
        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName);
        const range = sheet.getRange(1, 1, SheetUtils.getLastRow(sheet), SheetUtils.getLastColumn(sheet));
        this.reloadIssueData(range);
    }
}
class IssueHierarchyFormatter {
    static formatHierarchy(range) {
        if (![GSheetProjectSettings.childIssueColumnName].some(columnName => RangeUtils.doesRangeHaveSheetColumn(range, GSheetProjectSettings.sheetName, columnName))) {
            return;
        }
        let issuesRange = RangeUtils.toColumnRange(range, GSheetProjectSettings.issueColumnName);
        if (issuesRange == null) {
            return;
        }
        const sheet = issuesRange.getSheet();
        issuesRange = RangeUtils.withMinMaxRows(issuesRange, GSheetProjectSettings.firstDataRow, SheetUtils.getLastRow(sheet));
        const issues = Utils.timed(`${IssueHierarchyFormatter.name}: getting issues`, () => issuesRange.getValues()
            .map(it => { var _a; return (_a = it[0]) === null || _a === void 0 ? void 0 : _a.toString(); })
            .filter(it => it === null || it === void 0 ? void 0 : it.length)
            .filter(Utils.distinct()));
        if (!issues.length) {
            return;
        }
        if (GSheetProjectSettings.reorderHierarchyAutomatically) {
            Utils.timed(`${IssueHierarchyFormatter.name}: ${this.reorderIssuesAccordingToHierarchy.name}`, () => this.reorderIssuesAccordingToHierarchy(issues));
        }
        Utils.timed(`${IssueHierarchyFormatter.name}: ${this.formatHierarchyIssues.name}`, () => this.formatHierarchyIssues(issues));
    }
    static reorderAllIssuesAccordingToHierarchy() {
        this.reorderIssuesAccordingToHierarchy(undefined);
    }
    static reorderIssuesAccordingToHierarchy(issuesToReorder) {
        if (issuesToReorder != null && !issuesToReorder.length) {
            return;
        }
        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName);
        ProtectionLocks.lockAllRows(sheet);
        const issuesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueColumnName);
        const childIssuesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName);
        const { issues, childIssues, } = SheetUtils.getColumnsStringValues(sheet, {
            issues: issuesColumn,
            childIssues: childIssuesColumn,
        }, GSheetProjectSettings.firstDataRow);
        const notEmptyIssues = issues.filter(it => it === null || it === void 0 ? void 0 : it.length);
        const notEmptyUniqueIssues = notEmptyIssues.filter(Utils.distinct());
        Utils.trimArrayEndBy(issues, it => !(it === null || it === void 0 ? void 0 : it.length));
        SheetUtils.setLastRow(sheet, GSheetProjectSettings.firstDataRow + issues.length - 1);
        childIssues.length = issues.length;
        if (notEmptyIssues.length === notEmptyUniqueIssues.length) {
            return GSheetProjectSettings.firstDataRow + issues.length;
        }
        const moveIssues = (fromIndex, count, targetIndex) => {
            if (fromIndex === targetIndex || count <= 0) {
                return;
            }
            const fromRow = GSheetProjectSettings.firstDataRow + fromIndex;
            const targetRow = GSheetProjectSettings.firstDataRow + targetIndex;
            if (count === 1) {
                console.info(`Moving row #${fromRow} to #${targetRow}`);
            }
            else {
                console.info(`Moving rows #${fromRow}-${fromRow + count - 1} to #${targetRow}`);
            }
            const range = sheet.getRange(fromRow, 1, count, 1);
            sheet.moveRows(range, targetRow);
            Utils.moveArrayElements(issues, fromIndex, count, targetIndex);
            Utils.moveArrayElements(childIssues, fromIndex, count, targetIndex);
        };
        const groupIndexes = (indexes, targetIndex) => {
            while (indexes.length) {
                let index = indexes.shift();
                if (index === targetIndex) {
                    continue;
                }
                let bulkSize = 1;
                while (indexes.length) {
                    const nextIndex = indexes[0];
                    if (nextIndex === index + bulkSize) {
                        ++bulkSize;
                        indexes.shift();
                    }
                    else {
                        break;
                    }
                }
                moveIssues(index, bulkSize, targetIndex + 1);
                if (index < targetIndex) {
                    targetIndex += bulkSize - 1;
                }
                else {
                    targetIndex += bulkSize - (index < targetIndex ? 1 : 0);
                }
            }
        };
        const hasGapsInIndexes = (indexes) => {
            return indexes.length >= 2 && indexes[indexes.length - 1] - indexes[0] >= indexes.length;
        };
        for (const issue of notEmptyUniqueIssues) {
            if (issuesToReorder != null && !issuesToReorder.includes(issue)) {
                continue;
            }
            const getIndexes = () => issues
                .map((it, index) => issue === it ? index : null)
                .filter(index => index != null)
                .map(index => index)
                .toSorted(Utils.numericAsc());
            const getIndexesWithoutChild = () => getIndexes()
                .filter(index => { var _a; return !((_a = childIssues[index]) === null || _a === void 0 ? void 0 : _a.length); });
            const getIndexesWithChild = () => getIndexes()
                .filter(index => { var _a; return (_a = childIssues[index]) === null || _a === void 0 ? void 0 : _a.length; });
            { // group issues without child
                const indexesWithoutChild = getIndexesWithoutChild();
                if (hasGapsInIndexes(indexesWithoutChild)) {
                    const firstIndexWithoutChild = indexesWithoutChild.shift();
                    groupIndexes(indexesWithoutChild, firstIndexWithoutChild);
                }
            }
            { // group indexes with child
                const indexesWithChild = getIndexesWithChild();
                if (indexesWithChild.length) {
                    const indexesWithoutChild = getIndexesWithoutChild();
                    if (indexesWithoutChild.length) {
                        let targetIndex = getIndexesWithoutChild().pop();
                        if (indexesWithChild[0] >= targetIndex) {
                            ++targetIndex;
                        }
                        groupIndexes(indexesWithChild, targetIndex);
                    }
                    else if (hasGapsInIndexes(indexesWithChild)) {
                        const firstIndexWithChild = indexesWithChild.shift();
                        groupIndexes(indexesWithChild, firstIndexWithChild);
                    }
                }
            }
        }
    }
    static formatHierarchyIssues(issuesToFormat) {
        if (!issuesToFormat.length) {
            return;
        }
        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName);
        ProtectionLocks.lockAllRows(sheet);
        const issuesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueColumnName);
        const childIssuesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName);
        const milestonesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.milestoneColumnName);
        const typesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.typeColumnName);
        const deadlinesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.deadlineColumnName);
        const titlesColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.titleColumnName);
        const { issues, childIssues, milestones, types, deadlines, } = SheetUtils.getColumnsStringValues(sheet, {
            issues: issuesColumn,
            childIssues: childIssuesColumn,
            milestones: milestonesColumn,
            types: typesColumn,
            deadlines: deadlinesColumn,
        }, GSheetProjectSettings.firstDataRow);
        const notEmptyIssues = issues.filter(it => it === null || it === void 0 ? void 0 : it.length);
        const notEmptyUniqueIssues = notEmptyIssues.filter(Utils.distinct());
        if (notEmptyIssues.length === notEmptyUniqueIssues.length) {
            return;
        }
        Utils.trimArrayEndBy(issues, it => !(it === null || it === void 0 ? void 0 : it.length));
        SheetUtils.setLastRow(sheet, GSheetProjectSettings.firstDataRow + issues.length - 1);
        childIssues.length = issues.length;
        milestones.length = issues.length;
        types.length = issues.length;
        deadlines.length = issues.length;
        const { milestoneFormulas, typeFormulas, deadlineFormulas, } = SheetUtils.getColumnsFormulas(sheet, {
            milestoneFormulas: milestonesColumn,
            typeFormulas: typesColumn,
            deadlineFormulas: deadlinesColumn,
        }, GSheetProjectSettings.firstDataRow);
        milestoneFormulas.length = issues.length;
        typeFormulas.length = issues.length;
        deadlineFormulas.length = issues.length;
        for (const issue of notEmptyUniqueIssues) {
            if (!issuesToFormat.includes(issue)) {
                continue;
            }
            const getIndexes = () => issues
                .map((it, index) => issue === it ? index : null)
                .filter(index => index != null)
                .map(index => index)
                .toSorted(Utils.numericAsc());
            const getIndexesWithoutChild = () => getIndexes()
                .filter(index => { var _a; return !((_a = childIssues[index]) === null || _a === void 0 ? void 0 : _a.length); });
            const getIndexesWithChild = () => getIndexes()
                .filter(index => { var _a; return (_a = childIssues[index]) === null || _a === void 0 ? void 0 : _a.length; });
            { // set indent
                const setIndent = (indexes, indent) => {
                    if (!indexes.length) {
                        return;
                    }
                    const notations = indexes.flatMap(index => {
                        const row = GSheetProjectSettings.firstDataRow + index;
                        return [
                            sheet.getRange(row, issuesColumn).getA1Notation(),
                            sheet.getRange(row, titlesColumn).getA1Notation(),
                        ];
                    });
                    const numberFormat = indent > 0
                        ? ' '.repeat(indent) + '@'
                        : '@';
                    sheet.getRangeList(notations)
                        .setNumberFormat(numberFormat)
                        .setFontLine('none');
                };
                setIndent(getIndexesWithoutChild(), 0);
                setIndent(getIndexesWithChild(), GSheetProjectSettings.indent);
            }
            { // set formulas
                const indexesWithoutChild = getIndexesWithoutChild();
                const indexesWithChild = getIndexesWithChild();
                if (indexesWithoutChild.length && indexesWithChild.length) {
                    const firstIndexWithoutChild = indexesWithoutChild[0];
                    const firstRowWithoutChild = GSheetProjectSettings.firstDataRow + firstIndexWithoutChild;
                    const getIssueFormula = (column) => RangeUtils.getAbsoluteReferenceFormula(sheet.getRange(firstRowWithoutChild, column));
                    const firstIndexWithChild = indexesWithChild[0];
                    const firstRowWithChild = GSheetProjectSettings.firstDataRow + firstIndexWithChild;
                    sheet.getRange(firstRowWithChild, issuesColumn, indexesWithChild.length, 1)
                        .setFormula(getIssueFormula(issuesColumn));
                    indexesWithChild.forEach(index => {
                        var _a, _b, _c, _d, _e, _f;
                        const row = GSheetProjectSettings.firstDataRow + index;
                        if (!((_a = milestones[index]) === null || _a === void 0 ? void 0 : _a.length) && !((_b = milestoneFormulas[index]) === null || _b === void 0 ? void 0 : _b.length)) {
                            sheet.getRange(row, milestonesColumn)
                                .setFormula(getIssueFormula(milestonesColumn));
                        }
                        if (!((_c = types[index]) === null || _c === void 0 ? void 0 : _c.length) && !((_d = typeFormulas[index]) === null || _d === void 0 ? void 0 : _d.length)) {
                            sheet.getRange(row, typesColumn)
                                .setFormula(getIssueFormula(typesColumn));
                        }
                        if (!((_e = deadlines[index]) === null || _e === void 0 ? void 0 : _e.length) && !((_f = deadlineFormulas[index]) === null || _f === void 0 ? void 0 : _f.length)) {
                            sheet.getRange(row, deadlinesColumn)
                                .setFormula(getIssueFormula(deadlinesColumn));
                        }
                    });
                }
            }
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
            if (issueId === null || issueId === void 0 ? void 0 : issueId.length) {
                const canonizedKey = this.issueIdToIssueKey(issueId);
                if (canonizedKey === null || canonizedKey === void 0 ? void 0 : canonizedKey.length) {
                    return canonizedKey;
                }
            }
        }
        {
            const searchQuery = this.extractSearchQuery(issueKey);
            if (searchQuery === null || searchQuery === void 0 ? void 0 : searchQuery.length) {
                const canonizedKey = this.searchQueryToIssueKey(searchQuery);
                if (canonizedKey === null || canonizedKey === void 0 ? void 0 : canonizedKey.length) {
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
        if (!(issueIds === null || issueIds === void 0 ? void 0 : issueIds.length)) {
            return undefined;
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.getUrlForIssueIds.name);
    }
    loadIssues(issueIds) {
        if (!(issueIds === null || issueIds === void 0 ? void 0 : issueIds.length)) {
            return [];
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.loadIssues.name);
    }
    loadChildren(issueIds) {
        if (!(issueIds === null || issueIds === void 0 ? void 0 : issueIds.length)) {
            return [];
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.loadChildren.name);
    }
    loadBlockers(issueIds) {
        if (!(issueIds === null || issueIds === void 0 ? void 0 : issueIds.length)) {
            return [];
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.loadBlockers.name);
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
    search(query) {
        if (!(query === null || query === void 0 ? void 0 : query.length)) {
            return [];
        }
        throw Utils.throwNotImplemented(this.constructor.name, this.search.name);
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
    issueIdToIssueKey(issueId) {
        return `example/${issueId}`;
    }
    extractIssueId(issueKey) {
        return Utils.extractRegex(issueKey, /^example\/([\d.-]+)$/, 1);
    }
    getUrlForIssueId(issueId) {
        return `https://example.com/issues/${encodeURIComponent(issueId)}`;
    }
    getUrlForIssueIds(issueIds) {
        if (!(issueIds === null || issueIds === void 0 ? void 0 : issueIds.length)) {
            return null;
        }
        return `https://example.com/search?q=id:(${encodeURIComponent(issueIds.join('|'))})`;
    }
    loadIssues(issueIds) {
        if (!(issueIds === null || issueIds === void 0 ? void 0 : issueIds.length)) {
            return [];
        }
        return issueIds.map(id => new IssueExample(this, id));
    }
    loadChildren(issueIds) {
        if (!(issueIds === null || issueIds === void 0 ? void 0 : issueIds.length)) {
            return [];
        }
        return issueIds.flatMap(id => {
            let hash = parseInt(id);
            if (isNaN(hash)) {
                hash = Math.abs(Utils.hashCode(id));
            }
            return Array.from(Utils.range(0, hash % 3)).map(index => new IssueExample(this, `${id}-${index + 1}`));
        });
    }
    loadBlockers(issueIds) {
        if (!(issueIds === null || issueIds === void 0 ? void 0 : issueIds.length)) {
            return [];
        }
        return issueIds.flatMap(id => {
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
    search(query) {
        if (!(query === null || query === void 0 ? void 0 : query.length)) {
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
        const proxy = new Proxy({}, {
            apply(_, thisArg, argArray) {
                const instance = lazy.get();
                return Reflect.apply(instance, thisArg, argArray);
            },
            construct(_, argArray, newTarget) {
                const instance = lazy.get();
                return Reflect.construct(instance, argArray, newTarget);
            },
            defineProperty(_, property, attributes) {
                const instance = lazy.get();
                return Reflect.defineProperty(instance, property, attributes);
            },
            deleteProperty(_, property) {
                const instance = lazy.get();
                return Reflect.deleteProperty(instance, property);
            },
            get(_, property) {
                const instance = lazy.get();
                return Reflect.get(instance, property, instance);
            },
            getOwnPropertyDescriptor(_, property) {
                const instance = lazy.get();
                return Reflect.getOwnPropertyDescriptor(instance, property);
            },
            getPrototypeOf(_) {
                const instance = lazy.get();
                return Reflect.getPrototypeOf(instance);
            },
            has(_, property) {
                const instance = lazy.get();
                return Reflect.has(instance, property);
            },
            isExtensible(_) {
                const instance = lazy.get();
                return Reflect.isExtensible(instance);
            },
            ownKeys(_) {
                const instance = lazy.get();
                return Reflect.ownKeys(instance);
            },
            preventExtensions(_) {
                const instance = lazy.get();
                return Reflect.preventExtensions(instance);
            },
            set(_, property, newValue) {
                const instance = lazy.get();
                return Reflect.set(instance, property, newValue, instance);
            },
            setPrototypeOf(_, value) {
                const instance = lazy.get();
                return Reflect.setPrototypeOf(instance, value);
            },
        });
        return proxy;
    }
}
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
        var _a;
        return (_a = this.findNamedRange(rangeName)) !== null && _a !== void 0 ? _a : (() => {
            throw new Error(`"${rangeName}" named range can't be found`);
        })();
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
        Utils.timed(`${ProtectionLocks.name}: ${this.lockAllColumns.name}: ${sheet.getSheetName()}`, () => {
            const range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
            const protection = range.protect()
                .setDescription(`lock|columns|all|${new Date().getTime()}`)
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
        Utils.timed(`${ProtectionLocks.name}: ${this.lockAllRows.name}: ${sheet.getSheetName()}`, () => {
            const range = sheet.getRange(1, sheet.getMaxColumns(), sheet.getMaxRows(), 1);
            const protection = range.protect()
                .setDescription(`lock|rows|all|${new Date().getTime()}`)
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
            Utils.timed(`${ProtectionLocks.name}: ${this.lockRows.name}: ${sheet.getSheetName()}: ${rowToLock}`, () => {
                const range = sheet.getRange(1, sheet.getMaxColumns(), rowToLock, 1);
                const protection = range.protect()
                    .setDescription(`lock|rows|${rowToLock}|${new Date().getTime()}`)
                    .setWarningOnly(true);
                rowsProtections.set(rowToLock, protection);
            });
        }
    }
    static release() {
        if (!GSheetProjectSettings.lockColumns && !GSheetProjectSettings.lockRows) {
            return;
        }
        Utils.timed(`${ProtectionLocks.name}: ${this.release.name}`, () => {
            this._allColumnsProtections.forEach(protection => protection.remove());
            this._allColumnsProtections.clear();
            this._allRowsProtections.forEach(protection => protection.remove());
            this._allRowsProtections.clear();
            this._rowsProtections.forEach(protections => Array.from(protections.values()).forEach(protection => protection.remove()));
            this._rowsProtections.clear();
        });
    }
    static releaseExpiredLocks() {
        if (!GSheetProjectSettings.lockColumns && !GSheetProjectSettings.lockRows) {
            return;
        }
        Utils.timed(`${ProtectionLocks.name}: ${this.releaseExpiredLocks.name}`, () => {
            const maxLockDurationMillis = 10 * 60 * 1000;
            const minTimestamp = new Date().getTime() - maxLockDurationMillis;
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
            var _a, _b;
            const title = ((_a = link.title) === null || _a === void 0 ? void 0 : _a.length)
                ? link.title
                : link.url;
            if (!(title === null || title === void 0 ? void 0 : title.length)) {
                return;
            }
            if (text.length) {
                text += '\n';
            }
            if ((_b = link.url) === null || _b === void 0 ? void 0 : _b.length) {
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
        var _a;
        return `${((_a = this.constructor) === null || _a === void 0 ? void 0 : _a.name) || Utils.normalizeName(this.sheetName)}:migrate:`;
    }
    get _documentFlag() {
        return `${this._documentFlagPrefix}66728af16c370194fd38abbbd100b8404039e9527a8930d0b8bed31934955bd5:${GSheetProjectSettings.computeStringSettingsHash()}`;
    }
    migrateIfNeeded() {
        if (DocumentFlags.isSet(this._documentFlag)) {
            return;
        }
        this.migrate();
    }
    migrate() {
        var _a, _b, _c, _d, _e, _f;
        const sheet = this.sheet;
        ConditionalFormatting.removeConditionalFormatRulesByScope(sheet, 'layout');
        const columns = this.columns.reduce((map, info) => {
            map.set(Utils.normalizeName(info.name), info);
            return map;
        }, new Map());
        if (!columns.size) {
            return;
        }
        ProtectionLocks.lockAllColumns(sheet);
        const columnByKey = new Map();
        let lastColumn = SheetUtils.getLastColumn(sheet);
        const maxRows = sheet.getMaxRows();
        const existingNormalizedNames = sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn)
            .getValues()[0]
            .map(it => it === null || it === void 0 ? void 0 : it.toString())
            .map(it => (it === null || it === void 0 ? void 0 : it.length) ? Utils.normalizeName(it) : '');
        for (const [columnName, info] of columns.entries()) {
            const existingIndex = existingNormalizedNames.indexOf(columnName);
            if (existingIndex >= 0) {
                if ((_a = info.key) === null || _a === void 0 ? void 0 : _a.length) {
                    const columnNumber = existingIndex + 1;
                    columnByKey.set(info.key, { columnNumber, info });
                }
                continue;
            }
            console.info(`Adding "${info.name}" column`);
            ++lastColumn;
            const titleRange = sheet.getRange(GSheetProjectSettings.titleRow, lastColumn)
                .setValue(info.name);
            ExecutionCache.resetCache();
            if ((_b = info.key) === null || _b === void 0 ? void 0 : _b.length) {
                const columnNumber = lastColumn;
                columnByKey.set(info.key, { columnNumber, info });
            }
            if (info.defaultFontSize) {
                titleRange.setFontSize(info.defaultFontSize);
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
            if ((_c = info.defaultHorizontalAlignment) === null || _c === void 0 ? void 0 : _c.length) {
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
            if ((_d = info.arrayFormula) === null || _d === void 0 ? void 0 : _d.length) {
                const formulaToExpect = `
                    ={
                        "${Utils.escapeFormulaString(info.name)}";
                        ${Utils.processFormula(info.arrayFormula)}
                    }
                `;
                const formula = existingFormulas.get()[index];
                if (formula !== formulaToExpect) {
                    sheet.getRange(GSheetProjectSettings.titleRow, column)
                        .setFormula(formulaToExpect);
                }
            }
            const range = sheet.getRange(GSheetProjectSettings.firstDataRow, column, maxRows, 1);
            if ((_e = info.rangeName) === null || _e === void 0 ? void 0 : _e.length) {
                SpreadsheetApp.getActiveSpreadsheet().setNamedRange(info.rangeName, range);
            }
            const processFormula = (formula) => {
                formula = Utils.processFormula(formula);
                formula = formula.replaceAll(/#COLUMN_CELL\(([^)]+)\)/g, (_, key) => {
                    var _a;
                    const columnNumber = (_a = columnByKey.get(key)) === null || _a === void 0 ? void 0 : _a.columnNumber;
                    if (columnNumber == null) {
                        throw new Error(`Column with key '${key}' can't be found`);
                    }
                    return sheet.getRange(GSheetProjectSettings.firstDataRow, columnNumber).getA1Notation();
                });
                formula = formula.replaceAll(/#COLUMN_CELL\b/g, () => {
                    return range.getCell(1, 1).getA1Notation();
                });
                return formula;
            };
            let dataValidation = info.dataValidation != null
                ? info.dataValidation()
                : null;
            if (dataValidation != null) {
                if (dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CUSTOM_FORMULA) {
                    const formula = processFormula(dataValidation.getCriteriaValues()[0].toString());
                    dataValidation = dataValidation.copy()
                        .requireFormulaSatisfied(formula)
                        .build();
                }
            }
            range.setDataValidation(dataValidation);
            (_f = info.conditionalFormats) === null || _f === void 0 ? void 0 : _f.forEach(rule => {
                const originalConfigurer = rule.configurer;
                rule.configurer = builder => {
                    originalConfigurer(builder);
                    const formula = ConditionalFormatRuleUtils.extractFormula(builder);
                    if (formula != null) {
                        builder.whenFormulaSatisfied(processFormula(formula));
                    }
                    return builder;
                };
                const fullRule = {
                    scope: 'layout',
                    ...rule,
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
                defaultFontSize: 1,
                defaultWidth: '#default-height',
                defaultHorizontalAlignment: 'center',
            },
            
            {
                name: GSheetProjectSettings.milestoneColumnName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.typeColumnName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.issueColumnName,
                rangeName: GSheetProjectSettings.issuesRangeName,
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(`=OR(
                            ${GSheetProjectSettings.childIssuesRangeName} <> "",
                            COUNTIFS(
                                ${GSheetProjectSettings.issuesRangeName}, "=" & #SELF,
                                ${GSheetProjectSettings.childIssuesRangeName}, "="
                            ) <= 1
                        )`)
                    .setHelpText(`Multiple rows with the same ${GSheetProjectSettings.issueColumnName}`
                    + ` without ${GSheetProjectSettings.childIssueColumnName}`)
                    .build(),
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.childIssueColumnName,
                rangeName: GSheetProjectSettings.childIssuesRangeName,
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(`=COUNTIF(${GSheetProjectSettings.issuesRangeName}, "=" & #SELF) = 0`)
                    .setHelpText(`Only one level of hierarchy is supported`)
                    .build(),
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.titleColumnName,
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
                dataValidation: () => SpreadsheetApp.newDataValidation()
                    .requireFormulaSatisfied(`=INDIRECT(ADDRESS(ROW(), COLUMN(${GSheetProjectSettings.teamsRangeName}))) <> ""`)
                    .setHelpText(`Estimate must be defined for a team`)
                    .build(),
                defaultFormat: '#,##0',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.startColumnName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            {
                name: GSheetProjectSettings.endColumnName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
                conditionalFormats: [
                    {
                        order: 1,
                        configurer: builder => builder
                            .whenFormulaSatisfied(`=AND(
                                    ISFORMULA(#COLUMN_CELL),
                                    #COLUMN_CELL <> "",
                                    #COLUMN_CELL(deadline) <> "",
                                    #COLUMN_CELL > #COLUMN_CELL(deadline)
                                )`)
                            .setItalic(true)
                            .setBold(true)
                            .setFontColor('#c00'),
                    },
                    {
                        order: 2,
                        configurer: builder => builder
                            .whenFormulaSatisfied(`=AND(
                                    #COLUMN_CELL <> "",
                                    #COLUMN_CELL(deadline) <> "",
                                    #COLUMN_CELL > #COLUMN_CELL(deadline)
                                )`)
                            .setBold(true)
                            .setFontColor('#f00'),
                    },
                ],
            },
            {
                key: 'deadline',
                name: GSheetProjectSettings.deadlineColumnName,
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
        this.instances.forEach(instance => instance.migrateIfNeeded());
        CommonFormatter.applyCommonFormatsToAllSheets();
    }
    static migrate() {
        this.instances.forEach(instance => instance.migrate());
        CommonFormatter.applyCommonFormatsToAllSheets();
    }
}
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
        if (!(sheetName === null || sheetName === void 0 ? void 0 : sheetName.length)) {
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
        var _a;
        return (_a = this.findSheetByName(sheetName)) !== null && _a !== void 0 ? _a : (() => {
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
        ExecutionCache.put(['last-row', sheet], Math.max(lastRow, 1));
    }
    static getLastColumn(sheet) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        return ExecutionCache.getOrCompute(['last-column', sheet], () => Math.max(sheet.getLastColumn(), 1));
    }
    static setLastColumn(sheet, lastColumn) {
        ExecutionCache.put(['last-column', sheet], lastColumn);
    }
    static findColumnByName(sheet, columnName) {
        if (!(columnName === null || columnName === void 0 ? void 0 : columnName.length)) {
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
        var _a;
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        return (_a = this.findColumnByName(sheet, columnName)) !== null && _a !== void 0 ? _a : (() => {
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
        Utils.timed([
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
class Timer {
    constructor(name) {
        this._name = name;
        this._start = Date.now();
    }
    log() {
        console.log(`${this._name}: ${Date.now() - this._start}ms`);
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
    static processFormula(formula) {
        return formula
            .replaceAll(/#SELF\b/g, 'INDIRECT("RC", FALSE)')
            .split(/[\r\n]+/)
            .map(line => line.replace(/^\s+/, ''))
            .filter(line => line.length)
            .map(line => line.replaceAll(/^([*/+-]+ )/g, ' $1'))
            .map(line => line.replaceAll(/\s*\t\s*/g, ' '))
            .map(line => line.replaceAll(/"\s*&""/g, '"'))
            .map(line => line + (line.endsWith(',') || line.endsWith(';') ? ' ' : ''))
            .join('')
            .trim();
    }
    static escapeFormulaString(string) {
        return string.replaceAll(/"/g, '""');
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
    static extractRegex(string, regexp, group) {
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
        for (let index = 0; index <= array.length; ++index) {
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
        var _a;
        return (_a = this.parseDate(value)) !== null && _a !== void 0 ? _a : (() => {
            throw new Error(`Not a date: "${value}"`);
        })();
    }
    static hashCode(value) {
        if (!(value === null || value === void 0 ? void 0 : value.length)) {
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
        const array = new Array(length !== null && length !== void 0 ? length : 0);
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