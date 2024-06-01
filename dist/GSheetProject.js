class GSheetProject {
    static reloadIssues() {
        EntryPoint.entryPoint(() => {
        });
    }
    static migrateColumns() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateColumns();
        });
    }
    static cleanup() {
        EntryPoint.entryPoint(() => {
            ProtectionLocks.releaseExpiredLocks();
        });
    }
    static onOpen(event) {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateColumns();
        });
    }
    static onChange(event) {
        var _a;
        const changeType = (_a = event === null || event === void 0 ? void 0 : event.changeType) === null || _a === void 0 ? void 0 : _a.toString();
        if (changeType === 'INSERT_ROW') {
            this._onInsertRow();
        }
        else if (changeType === 'REMOVE_COLUMN') {
            this._onRemoveColumn();
        }
    }
    static _onInsertRow() {
        EntryPoint.entryPoint(() => {
            IssueHierarchyFormatter.formatHierarchyForAllIssues();
        });
    }
    static _onRemoveColumn() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateColumns();
        });
    }
    static onEdit(event) {
        this._onEditRange(event === null || event === void 0 ? void 0 : event.range);
    }
    static onFormSubmit(event) {
        this._onEditRange(event === null || event === void 0 ? void 0 : event.range);
    }
    static _onEditRange(range) {
        if (range == null) {
            return;
        }
        EntryPoint.entryPoint(() => {
            IssueHierarchyFormatter.formatHierarchy(range);
        });
    }
}
class GSheetProjectSettings {
    static computeStringSettingsHash() {
        const hashableValues = {};
        for (const [key, value] of Object.entries(GSheetProjectSettings)) {
            if (Utils.isString(value)) {
                hashableValues[key] = value;
            }
        }
        const json = JSON.stringify(hashableValues);
        return SHA256(json);
    }
}
GSheetProjectSettings.titleRow = 1;
GSheetProjectSettings.firstDataRow = 2;
GSheetProjectSettings.settingsSheetName = "Settings";
GSheetProjectSettings.projectsSheetName = "Projects";
GSheetProjectSettings.projectsIconColumnName = "Icon";
GSheetProjectSettings.projectsDoneColumnName = "Done";
GSheetProjectSettings.projectsIssueColumnName = "Issue";
//static projectsIssuesRangeName: string = "Issues"
GSheetProjectSettings.projectsParentIssueColumnName = "Parent Issue";
GSheetProjectSettings.projectsTitleColumnName = "Title";
GSheetProjectSettings.projectsTeamColumnName = "Team";
GSheetProjectSettings.projectsEstimateColumnName = "Estimate (days)";
GSheetProjectSettings.projectsDeadlineColumnName = "Deadline";
GSheetProjectSettings.projectsStartColumnName = "Start";
GSheetProjectSettings.projectsEndColumnName = "End";
//static projectsIssueHashColumnName: string = "Issue Hash"
GSheetProjectSettings.indent = 4;
GSheetProjectSettings.issueLoaderFactories = [];
GSheetProjectSettings.issueChildrenLoaderFactories = [];
GSheetProjectSettings.issueBlockersLoaderFactories = [];
GSheetProjectSettings.issueSearcherFactories = [];
class CommonFormatter {
    static setMiddleVerticalAlign() {
        SpreadsheetApp.getActiveSpreadsheet().getSheets()
            .filter(sheet => SheetUtils.isGridSheet(sheet))
            .forEach(sheet => {
            sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
                .setVerticalAlignment('middle');
        });
    }
    static onChange(event) {
        var _a, _b;
        if (['INSERT_ROW', 'INSERT_COLUMN'].includes((_b = (_a = event === null || event === void 0 ? void 0 : event.changeType) === null || _a === void 0 ? void 0 : _a.toString()) !== null && _b !== void 0 ? _b : '')) {
            this.setMiddleVerticalAlign();
        }
    }
}
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
        try {
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
        }
    }
}
class ExecutionCache {
    static getOrComputeCache(key, compute) {
        const stringKey = JSON.stringify(key, (_, value) => {
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
        if (this._data.has(stringKey)) {
            return this._data.get(stringKey);
        }
        const result = compute();
        this._data.set(stringKey, result);
        return result;
    }
    static resetCache() {
        this._data.clear();
    }
}
ExecutionCache._data = new Map();
class Images {
}
Images.loadingImageUrl = 'https://raw.githubusercontent.com/remal/misc/main/spinner.gif';
class IssueBlockersLoader {
    loadBlockers(issueId) {
        return this.loadBlockersBulk([issueId]);
    }
    loadBlockersBulk(issueIds) {
        return [];
    }
}
class IssueBlockersLoaderFactory {
    getIssueBlockerLoader(issueId) {
        return undefined;
    }
}
class IssueChildrenLoader {
    loadChildren(issueId) {
        return this.loadChildrenBulk([issueId]);
    }
    loadChildrenBulk(issueIds) {
        return [];
    }
}
class IssueChildrenLoaderFactory {
    getIssueChildrenLoader(issueId) {
        return undefined;
    }
}
class IssueHierarchyFormatter {
    static formatHierarchy(range) {
        const issues = [];
        const issuesRange = RangeUtils.toColumnRange(range, GSheetProjectSettings.projectsIssueColumnName);
        if (issuesRange != null) {
            issuesRange.getValues()
                .map(it => { var _a; return (_a = it[0]) === null || _a === void 0 ? void 0 : _a.toString(); })
                .forEach(it => issues.push(it));
        }
        const parentIssuesRange = RangeUtils.toColumnRange(range, GSheetProjectSettings.projectsParentIssueColumnName);
        if (parentIssuesRange != null) {
            parentIssuesRange.getValues()
                .map(it => { var _a; return (_a = it[0]) === null || _a === void 0 ? void 0 : _a.toString(); })
                .forEach(it => issues.push(it));
        }
        this.formatHierarchyForIssues(issues);
    }
    static formatHierarchyForAllIssues() {
        const issues = [];
        const parentIssuesRange = SheetUtils.getColumnRange(GSheetProjectSettings.projectsSheetName, GSheetProjectSettings.projectsParentIssueColumnName, GSheetProjectSettings.firstDataRow);
        parentIssuesRange.getValues()
            .map(it => { var _a; return (_a = it[0]) === null || _a === void 0 ? void 0 : _a.toString(); })
            .forEach(it => issues.push(it));
        this.formatHierarchyForIssues(issues);
    }
    static formatHierarchyForIssues(issues) {
        issues = issues
            .filter(it => it === null || it === void 0 ? void 0 : it.length)
            .filter(Utils.distinct());
        console.info('issues', issues);
        issues.forEach(issue => this.formatHierarchyForIssue(issue));
    }
    static formatHierarchyForIssue(issue) {
        var _a, _b, _c;
        console.info('issue', issue);
        const issueSlug = issue.replaceAll(/[\r\n]+/g, '').replace(/^(.{0,25}).*$/, '$1');
        console.info('issueSlug', issueSlug);
        const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.projectsSheetName);
        ProtectionLocks.lockRowsWithProtection(sheet);
        const issueRange = SheetUtils.getColumnRange(GSheetProjectSettings.projectsSheetName, GSheetProjectSettings.projectsIssueColumnName, GSheetProjectSettings.firstDataRow)
            .createTextFinder(issue)
            .ignoreDiacritics(false)
            .matchCase(true)
            .matchEntireCell(true)
            .findNext();
        if (issueRange == null) {
            return;
        }
        let issueRow = issueRange.getRow();
        console.info('issueRow', issueRow);
        const issueTitleRange = sheet.getRange(issueRow, SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsTitleColumnName));
        let indentLevel = Math.ceil(RangeUtils.getIndent(issueTitleRange) / GSheetProjectSettings.indent);
        console.info('indentLevel', indentLevel);
        const shouldIssueHaveIndent = (_c = (_b = (_a = sheet.getRange(issueRow, SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsParentIssueColumnName)).getValue()) === null || _a === void 0 ? void 0 : _a.toString()) === null || _b === void 0 ? void 0 : _b.trim()) === null || _c === void 0 ? void 0 : _c.length;
        console.info('shouldIssueHaveIndent', shouldIssueHaveIndent);
        if (!shouldIssueHaveIndent && indentLevel > 0) {
            indentLevel = 0;
            RangeUtils.setStringIndent(issueTitleRange, 0);
            console.info('indentLevel', indentLevel);
        }
        const childIssueRows = SheetUtils.getColumnRange(GSheetProjectSettings.projectsSheetName, GSheetProjectSettings.projectsParentIssueColumnName, GSheetProjectSettings.firstDataRow)
            .createTextFinder(issue)
            .ignoreDiacritics(false)
            .matchCase(true)
            .matchEntireCell(true)
            .findAll()
            .map(it => it.getRow())
            .filter(it => it !== issueRow);
        if (!childIssueRows.length) {
            return;
        }
        Utils.timed(`${IssueHierarchyFormatter.name}: ${issueSlug}: Adjust groups`, () => {
            for (const row of childIssueRows) {
                const currentGroupDepth = sheet.getRowGroupDepth(row);
                const expectedGroupDepth = Math.min(indentLevel + 1, 4);
                if (currentGroupDepth !== expectedGroupDepth) {
                    sheet.getRange(row, 1).shiftRowGroupDepth(expectedGroupDepth - currentGroupDepth);
                }
            }
        });
        const childIssueRanges = [];
        for (let rowIndex = 0; rowIndex < childIssueRows.length; ++rowIndex) {
            const row = childIssueRows[rowIndex];
            let rows = 1;
            let lastRow = row;
            console.info('lastRow', lastRow);
            for (let i = rowIndex + 1; i < childIssueRows.length; ++i) {
                const nextRow = childIssueRows[i];
                console.info('nextRow', nextRow);
                if (nextRow === lastRow + 1) {
                    ++rows;
                    lastRow = nextRow;
                    console.info('lastRow', lastRow);
                }
                else {
                    break;
                }
            }
            console.info('rows', rows);
            rowIndex += rows - 1;
            const combinedRange = sheet.getRange(row, 1, rows, 1);
            childIssueRanges.push(combinedRange);
        }
        console.info('childIssueRanges', childIssueRanges.map(range => `${range.getRow()}+${range.getNumRows() - 1}`));
        Utils.timed(`${IssueHierarchyFormatter.name}: ${issueSlug}: Adjust indents`, () => {
            for (const childIssueRange of childIssueRanges) {
                const childIssueTitleRange = sheet.getRange(childIssueRange.getRow(), SheetUtils.getColumnByName(sheet, GSheetProjectSettings.projectsTitleColumnName), childIssueRange.getNumRows(), 1);
                RangeUtils.setStringIndent(childIssueTitleRange, (indentLevel + 1) * GSheetProjectSettings.indent);
            }
        });
        // move children after the issue:
        Utils.timed(`${IssueHierarchyFormatter.name}: ${issueSlug}: Move children after the issue`, () => {
            let lastIssueOrConnectedChildIssueRow = issueRow;
            for (const childIssueRange of childIssueRanges) {
                const childIssueRow = childIssueRange.getRow();
                if (childIssueRow === issueRow + 1) {
                    lastIssueOrConnectedChildIssueRow += childIssueRange.getNumRows();
                    break;
                }
            }
            console.info('lastIssueOrConnectedChildIssueRow', lastIssueOrConnectedChildIssueRow);
            for (const childIssueRange of childIssueRanges) {
                const childIssueRow = childIssueRange.getRow();
                if (childIssueRow < issueRow) {
                    continue;
                }
                if (childIssueRow < lastIssueOrConnectedChildIssueRow) {
                    continue;
                }
                console.info('childIssueRow', childIssueRow);
                sheet.moveRows(childIssueRange, lastIssueOrConnectedChildIssueRow + 1);
                lastIssueOrConnectedChildIssueRow += childIssueRange.getNumRows();
                console.info('lastIssueOrConnectedChildIssueRow', lastIssueOrConnectedChildIssueRow);
            }
        });
        // move children before the issue:
        Utils.timed(`${IssueHierarchyFormatter.name}: ${issueSlug}: Move children before the issue`, () => {
            for (const childIssueRange of childIssueRanges.toReversed()) {
                const childIssueRow = childIssueRange.getRow();
                if (childIssueRow >= issueRow) {
                    continue;
                }
                console.info('childIssueRow', childIssueRow);
                sheet.moveRows(childIssueRange, issueRow + 1);
                issueRow -= childIssueRange.getNumRows();
                console.info('issueRow', issueRow);
            }
        });
    }
}
class IssueLoader {
    load(issueId) {
        return null;
    }
    canonizeId(issueId) {
        return issueId;
    }
    createWebUrl(issueId) {
        return null;
    }
}
class IssueLoaderFactory {
    getIssueLoader(issueId) {
        return undefined;
    }
}
class IssueSearcher {
    search(query) {
        return [];
    }
    canonizeQuery(query) {
        return query;
    }
    createWebUrl(query) {
        return null;
    }
}
class IssueSearcherFactory {
    getIssueSearcher(query) {
        return undefined;
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
class NamedRangeUtils {
    static findNamedRange(rangeName) {
        rangeName = Utils.normalizeName(rangeName);
        return ExecutionCache.getOrComputeCache(['findNamedRange', rangeName], () => {
            for (const namedRange of SpreadsheetApp.getActiveSpreadsheet().getNamedRanges()) {
                const name = Utils.normalizeName(namedRange.getName());
                if (name === rangeName) {
                    return namedRange;
                }
            }
            return undefined;
        });
    }
    static getNamedRange(rangeName) {
        var _a;
        return (_a = this.findNamedRange(rangeName)) !== null && _a !== void 0 ? _a : (() => {
            throw new Error(`"${rangeName}" named range can't be found`);
        })();
    }
}
class ProtectionLocks {
    static lockColumnsWithProtection(sheet) {
        const sheetId = sheet.getSheetId();
        if (this._columnsProtections.has(sheetId)) {
            return;
        }
        const range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
        const protection = range.protect()
            .setDescription(`lock|columns|${new Date().getTime()}`)
            .setWarningOnly(true);
        this._columnsProtections.set(sheetId, protection);
    }
    static lockRowsWithProtection(sheet) {
        const sheetId = sheet.getSheetId();
        if (this._rowsProtections.has(sheetId)) {
            return;
        }
        const range = sheet.getRange(1, sheet.getMaxColumns(), sheet.getMaxRows(), 1);
        const protection = range.protect()
            .setDescription(`lock|rows|${new Date().getTime()}`)
            .setWarningOnly(true);
        this._rowsProtections.set(sheetId, protection);
    }
    static release() {
        this._columnsProtections.forEach(protection => protection.remove());
        this._columnsProtections.clear();
        this._rowsProtections.forEach(protection => protection.remove());
        this._rowsProtections.clear();
    }
    static releaseExpiredLocks() {
        const maxLockDurationMillis = 10 * 60 * 1000;
        const minTimestamp = new Date().getTime() - maxLockDurationMillis;
        SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => {
            for (const protection of sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)) {
                const description = protection.getDescription();
                if (!description.startsWith('lock|')) {
                    continue;
                }
                const dateString = description.split('|').slice(-1)[0];
                try {
                    const date = Number.isNaN(dateString)
                        ? new Date(dateString)
                        : new Date(parseFloat(dateString));
                    if (date.getTime() < minTimestamp) {
                        console.warn(`Removing expired protection lock: ${description}`);
                        protection.remove();
                    }
                }
                catch (_) {
                    // do nothing
                }
            }
        });
    }
}
ProtectionLocks._columnsProtections = new Map();
ProtectionLocks._rowsProtections = new Map();
class RangeUtils {
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
        if (!this.doesRangeHaveColumn(range, column)) {
            return undefined;
        }
        return range.offset(0, column - range.getColumn(), range.getNumRows(), 1);
    }
    static doesRangeHaveColumn(range, column) {
        if (range == null || column == null) {
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
    static createLinksValue(links) {
        let text = '';
        const linksWithOffsets = [];
        links.forEach(link => {
            var _a;
            if (text.length) {
                text += '\n';
            }
            if (!((_a = link.title) === null || _a === void 0 ? void 0 : _a.length)) {
                link.title = link.url;
            }
            linksWithOffsets.push({
                url: link.url,
                start: text.length,
                end: text.length + link.title.length,
            });
            text += link.title;
        });
        const builder = SpreadsheetApp.newRichTextValue().setText(text);
        linksWithOffsets.forEach(link => builder.setLinkUrl(link.start, link.end, link.url));
        return builder.build();
    }
}
class Settings {
    static get settingsSheet() {
        return SheetUtils.getSheetByName(GSheetProjectSettings.settingsSheetName);
    }
    static getMatrix(settingsScope) {
        const settingsSheet = this.settingsSheet;
        settingsScope = Utils.normalizeName(settingsScope);
        return ExecutionCache.getOrComputeCache(['settings', 'matrix', settingsScope], () => {
            const scopeRow = this._findScopeRow(settingsSheet, settingsScope);
            if (scopeRow == null) {
                throw new Error(`Settings with "${settingsScope}" can't be found`);
            }
            const columns = [];
            const columnsValues = settingsSheet
                .getRange(scopeRow + 1, 1, 1, settingsSheet.getLastColumn())
                .getValues()[0];
            for (const column of columnsValues) {
                const name = Utils.toLowerCamelCase(column.toString().trim());
                if (name.length) {
                    columns.push(name);
                }
                else {
                    break;
                }
            }
            if (!columns.length) {
                return [];
            }
            const result = [];
            const allSettingsRange = result[':settingsRange'] = {
                row: scopeRow + 2,
                column: 1,
                rows: 0,
                columns: columns.length,
            };
            for (const row of Utils.range(scopeRow + 2, settingsSheet.getLastRow())) {
                const item = new Map();
                item[':settingsRange'] = {
                    row: row,
                    column: 1,
                    rows: 1,
                    columns: columns.length,
                };
                const values = settingsSheet.getRange(row, 1, 1, columns.length).getValues()[0];
                for (let i = 0; i < columns.length; ++i) {
                    let value = values[i].toString().trim();
                    item.set(columns[i], value);
                }
                const areAllValuesEmpty = Array.from(item.values()).every(value => !value.length);
                if (areAllValuesEmpty) {
                    break;
                }
                result.push(item);
                ++allSettingsRange.rows;
            }
            return result;
        });
    }
    static getMap(settingsScope) {
        const settingsSheet = this.settingsSheet;
        settingsScope = Utils.normalizeName(settingsScope);
        return ExecutionCache.getOrComputeCache(['settings', 'map', settingsScope], () => {
            const scopeRow = this._findScopeRow(settingsSheet, settingsScope);
            if (scopeRow == null) {
                throw new Error(`Settings with "${settingsScope}" can't be found`);
            }
            const result = new Map();
            const allSettingsRange = result[':settingsRange'] = {
                row: scopeRow + 1,
                column: 1,
                rows: 0,
                columns: 2,
            };
            for (const row of Utils.range(scopeRow + 1, settingsSheet.getLastRow())) {
                const values = settingsSheet.getRange(row, 1, 1, 2).getValues()[0];
                const key = Utils.toLowerCamelCase(values[0].toString().trim());
                if (!key.length) {
                    break;
                }
                const value = values[1].toString().trim();
                result.set(key, value);
                ++allSettingsRange.rows;
            }
            return result;
        });
    }
    static _findScopeRow(sheet, scope) {
        for (const row of Utils.range(1, sheet.getLastRow())) {
            const range = sheet.getRange(row, 1);
            if (range.getFontWeight() !== 'bold'
                || !range.isPartOfMerge()) {
                continue;
            }
            if (Utils.normalizeName(range.getValue()) === scope) {
                return row;
            }
        }
        return null;
    }
}
class SheetLayout {
    get sheet() {
        return SheetUtils.getSheetByName(this.sheetName);
    }
    migrateColumns() {
        var _a, _b, _c;
        const columns = this.columns.reduce((map, info) => {
            map.set(Utils.normalizeName(info.name), info);
            return map;
        }, new Map());
        if (!columns.size) {
            return;
        }
        const documentFlagPrefix = `${((_a = this.constructor) === null || _a === void 0 ? void 0 : _a.name) || Utils.normalizeName(this.sheetName)}:migrateColumns:`;
        const documentFlag = `${documentFlagPrefix}c24705e7d863de516fa39451144520ca21e8c6adf18e6ca9731fee1535301eab:${GSheetProjectSettings.computeStringSettingsHash()}`;
        if (DocumentFlags.isSet(documentFlag)) {
            return;
        }
        const sheet = this.sheet;
        ProtectionLocks.lockColumnsWithProtection(sheet);
        let lastColumn = Math.max(sheet.getLastColumn(), 1);
        const maxRows = sheet.getMaxRows();
        const existingNormalizedNames = sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn)
            .getValues()[0]
            .map(it => it === null || it === void 0 ? void 0 : it.toString())
            .map(it => (it === null || it === void 0 ? void 0 : it.length) ? Utils.normalizeName(it) : '');
        for (const [columnName, info] of columns.entries()) {
            if (!existingNormalizedNames.includes(columnName)) {
                const titleRange = sheet.getRange(GSheetProjectSettings.titleRow, lastColumn)
                    .setValue(info.name);
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
                existingNormalizedNames.push(columnName);
                ++lastColumn;
            }
        }
        const existingFormulas = new Lazy(() => sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn).getFormulas()[0]);
        for (const [columnName, info] of columns.entries()) {
            const index = existingNormalizedNames.indexOf(columnName);
            if (index < 0) {
                continue;
            }
            const column = index + 1;
            if ((_b = info.arrayFormula) === null || _b === void 0 ? void 0 : _b.length) {
                const arrayFormulaNormalized = info.arrayFormula.split(/[\r\n]+/)
                    .map(line => line.trim())
                    .filter(line => line.length)
                    .map(line => line + (line.endsWith(',') || line.endsWith(';') ? ' ' : ''))
                    .join('');
                const formulaToExpect = `={"${Utils.escapeFormulaString(info.name)}"; ${arrayFormulaNormalized}}`;
                const formula = existingFormulas.get()[index];
                if (formula !== formulaToExpect) {
                    sheet.getRange(GSheetProjectSettings.titleRow, column)
                        .setFormula(formulaToExpect);
                }
            }
            if ((_c = info.rangeName) === null || _c === void 0 ? void 0 : _c.length) {
                SpreadsheetApp.getActiveSpreadsheet().setNamedRange(info.rangeName, sheet.getRange(GSheetProjectSettings.firstDataRow, column, maxRows, 1));
            }
        }
        DocumentFlags.set(documentFlag);
        DocumentFlags.cleanupByPrefix(documentFlagPrefix);
        const waitForAllDataExecutionsCompletion = SpreadsheetApp.getActiveSpreadsheet()['waitForAllDataExecutionsCompletion'];
        if (Utils.isFunction(waitForAllDataExecutionsCompletion)) {
            waitForAllDataExecutionsCompletion(10);
        }
    }
}
class SheetLayoutProjects extends SheetLayout {
    get sheetName() {
        return GSheetProjectSettings.projectsSheetName;
    }
    get columns() {
        return [
            {
                name: GSheetProjectSettings.projectsIconColumnName,
                defaultFontSize: 1,
                defaultWidth: '#default-height',
            },
            {
                name: GSheetProjectSettings.projectsDoneColumnName,
            },
            {
                name: GSheetProjectSettings.projectsParentIssueColumnName,
            },
            {
                name: GSheetProjectSettings.projectsIssueColumnName,
                //rangeName: GSheetProjectSettings.projectsIssuesRangeName,
            },
            {
                name: GSheetProjectSettings.projectsTitleColumnName,
            },
            {
                name: GSheetProjectSettings.projectsTeamColumnName,
            },
            {
                name: GSheetProjectSettings.projectsEstimateColumnName,
            },
            {
                name: GSheetProjectSettings.projectsDeadlineColumnName,
            },
            {
                name: GSheetProjectSettings.projectsStartColumnName,
            },
            {
                name: GSheetProjectSettings.projectsEndColumnName,
            },
            /*
            {
                name: GSheetProjectSettings.projectsIssueHashColumnName,
                arrayFormula: `
                    MAP(
                        ARRAYFORMULA(${GSheetProjectSettings.projectsIssuesRangeName}),
                        LAMBDA(issue, IF(ISBLANK(issue), "", ${SHA256.name}(issue)))
                    )
                `,
            },
            */
        ];
    }
}
SheetLayoutProjects.instance = new SheetLayoutProjects();
class SheetLayouts {
    static migrateColumns() {
        this.instances.forEach(instance => instance.migrateColumns());
    }
}
SheetLayouts.instances = [
    SheetLayoutProjects.instance,
];
class SheetUtils {
    static findSheetByName(sheetName) {
        if (!(sheetName === null || sheetName === void 0 ? void 0 : sheetName.length)) {
            return undefined;
        }
        sheetName = Utils.normalizeName(sheetName);
        return ExecutionCache.getOrComputeCache(['findSheetByName', sheetName], () => {
            for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
                const name = Utils.normalizeName(sheet.getSheetName());
                if (name === sheetName) {
                    return sheet;
                }
            }
            return undefined;
        });
    }
    static getSheetByName(sheetName) {
        var _a;
        return (_a = this.findSheetByName(sheetName)) !== null && _a !== void 0 ? _a : (() => {
            throw new Error(`"${sheetName}" sheet can't be found`);
        })();
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
        ProtectionLocks.lockColumnsWithProtection(sheet);
        columnName = Utils.normalizeName(columnName);
        return ExecutionCache.getOrComputeCache(['findColumnByName', sheet, columnName], () => {
            for (const col of Utils.range(GSheetProjectSettings.titleRow, sheet.getLastColumn())) {
                const name = Utils.normalizeName(sheet.getRange(1, col).getValue());
                if (name === columnName) {
                    return col;
                }
            }
            return undefined;
        });
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
    static getColumnRange(sheet, column, fromRow) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        if (Utils.isString(column)) {
            column = this.getColumnByName(sheet, column);
        }
        if (fromRow == null) {
            fromRow = 1;
        }
        const lastRow = sheet.getLastRow();
        if (fromRow > lastRow) {
            return sheet.getRange(fromRow, column);
        }
        const rows = lastRow - fromRow + 1;
        return sheet.getRange(fromRow, column, rows, 1);
    }
    static getRowRange(sheet, row, fromColumn) {
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        if (fromColumn == null) {
            fromColumn = 1;
        }
        else if (Utils.isString(fromColumn)) {
            fromColumn = this.getColumnByName(sheet, fromColumn);
        }
        const lastColumn = sheet.getLastColumn();
        if (fromColumn > lastColumn) {
            return sheet.getRange(row, fromColumn);
        }
        const columns = lastColumn - fromColumn + 1;
        return sheet.getRange(row, fromColumn, 1, columns);
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
        const array = new Array(length);
        if (initValue !== undefined) {
            array.fill(initValue);
        }
        return array;
    }
    static escapeRegex(string) {
        return string.replaceAll(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    static escapeFormulaString(string) {
        return string.replaceAll(/"/g, '""');
    }
    static timed(timerLabel, action) {
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
    static isFunction(value) {
        return typeof value === 'function';
    }
    static isRecord(value) {
        return typeof value === 'object' && !Array.isArray(value);
    }
    static throwNotConfigured(name) {
        throw new Error(`Not configured: ${name}`);
    }
}