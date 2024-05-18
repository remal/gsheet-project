class GSheetProject {
    constructor(settings) {
        this.issueIdFormatter = new IssueIdFormatter(settings);
        this.issueInfoLoader = new IssueInfoLoader(settings);
    }
    onOpen(event) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache();
        });
    }
    onChange(event) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache();
        });
    }
    osEdit(event) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache();
            this.issueIdFormatter.formatIssueId(event.range);
            this.issueInfoLoader.loadIssueInfo(event.range);
        });
    }
    refresh() {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache();
            this.issueInfoLoader.loadAllIssueInfo();
        });
    }
}
class GSheetProjectSettings {
    constructor() {
        this.settingsSheetName = "Settings";
        this.issueIdColumnName = "Issue";
        this.parentIssueIdColumnName = "Parent Issue";
        this.idDoneCalculator = () => {
            throw new Error('idDoneCalculator is not set');
        };
        this.stringFields = {};
        this.booleanFields = {};
        this.childIssueMetrics = [];
        this.blockerIssueMetrics = [];
        this.issueIdsExtractor = () => {
            throw new Error('issueIdsExtractor is not set');
        };
        this.issueIdDecorator = (id) => id;
        this.issueIdToUrl = () => {
            throw new Error('issueIdToUrl is not set');
        };
        this.issueIdsToUrl = null;
        this.issuesLoader = () => {
            throw new Error('issuesLoader is not set');
        };
        this.childIssuesLoader = () => {
            throw new Error('childIssuesLoader is not set');
        };
        this.blockerIssuesLoader = () => {
            throw new Error('blockerIssuesLoader is not set');
        };
        this.issueIdGetter = () => {
            throw new Error('issueIdGetter is not set');
        };
    }
}
const DATA_FIRST_ROW = 2;
class ExecutionCache {
    static getOrComputeCache(key, compute) {
        const stringKey = JSON.stringify(key, (_, value) => {
            if (Utils.isFunction(value.getId)) {
                return value.getId();
            }
            else if (Utils.isFunction(value.getSheetId)) {
                return value.getSheetId();
            }
            return value;
        });
        if (this.data.has(stringKey)) {
            return this.data[stringKey];
        }
        const result = compute();
        this.data.set(stringKey, result);
        return result;
    }
    static resetCache() {
        this.data.clear();
    }
}
ExecutionCache.data = new Map();
class IssueIdFormatter {
    constructor(settings) {
        this.settings = settings;
    }
    formatIssueId(range) {
        const columnNames = [
            this.settings.issueIdColumnName,
            this.settings.parentIssueIdColumnName,
        ];
        for (const y of Utils.range(1, range.getHeight())) {
            for (const x of Utils.range(1, range.getWidth())) {
                const cell = range.getCell(y, x);
                if (!columnNames.some(name => RangeUtils.doesRangeHaveColumn(cell, name))) {
                    continue;
                }
                const ids = this.settings.issueIdsExtractor(cell.getValue());
                const links = ids.map(id => {
                    return {
                        url: this.settings.issueIdToUrl(id),
                        title: this.settings.issueIdDecorator(id),
                    };
                });
                cell.setValue(RichTextUtils.createLinksValue(links));
            }
        }
    }
}
class IssueInfoLoader {
    constructor(settings) {
        this.settings = settings;
    }
    loadIssueInfo(range) {
        if (!RangeUtils.doesRangeHaveColumn(range, this.settings.issueIdColumnName)) {
            return;
        }
        const sheet = range.getSheet();
        const rows = Array.from(Utils.range(1, range.getHeight()))
            .map(y => range.getCell(y, 1).getRow())
            .filter(row => row >= DATA_FIRST_ROW)
            .filter(Utils.distinct);
        for (const row of rows) {
            this.loadIssueInfoForRow(sheet, row);
        }
    }
    loadAllIssueInfo() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            const hasIssueIdColumn = SheetUtils.findColumnByName(sheet, this.settings.issueIdColumnName) != null;
            if (!hasIssueIdColumn) {
                return;
            }
            for (const row of Utils.range(DATA_FIRST_ROW, sheet.getLastRow())) {
                this.loadIssueInfoForRow(sheet, row);
            }
        }
    }
    loadIssueInfoForRow(sheet, row) {
        if (row < DATA_FIRST_ROW
            || sheet.isRowHiddenByUser(row)) {
            return;
        }
        const issueIdColumn = SheetUtils.getColumnByName(sheet, this.settings.issueIdColumnName);
        const issueIdRange = sheet.getRange(row, issueIdColumn);
        const issueIds = this.settings.issueIdsExtractor(issueIdRange.getValue());
        if (!issueIds.length) {
            return;
        }
        console.log(`"${sheet.getSheetName()}" sheet: processing row #${row}`);
        issueIdRange.setBackground('#eee');
        try {
            const rootIssues = this.settings.issuesLoader(issueIds);
            const childIssues = new Lazy(() => this.settings.childIssuesLoader(issueIds)
                .filter(issue => !issueIds.includes(this.settings.issueIdGetter(issue))));
            const blockerIssues = new Lazy(() => this.settings.blockerIssuesLoader(rootIssues.concat(childIssues.get())
                .map(issue => this.settings.issueIdGetter(issue))));
            const isDoneColumn = SheetUtils.findColumnByName(sheet, this.settings.isDoneColumnName);
            if (isDoneColumn != null) {
                const isDone = this.settings.idDoneCalculator(rootIssues, childIssues.get());
                sheet.getRange(row, isDoneColumn).setValue(isDone ? 'Yes' : '');
            }
            for (const [columnName, getter] of Object.entries(this.settings.stringFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName);
                if (fieldColumn != null) {
                    sheet.getRange(row, fieldColumn).setValue(rootIssues
                        .map(getter)
                        .join('\n'));
                }
            }
            for (const [columnName, getter] of Object.entries(this.settings.booleanFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName);
                if (fieldColumn != null) {
                    const isTrue = rootIssues.some(getter);
                    sheet.getRange(row, fieldColumn).setValue(isTrue ? 'Yes' : '');
                }
            }
            const calculateIssueMetrics = (metricsIssues, metrics) => {
                var _a;
                for (const metric of metrics) {
                    const metricColumn = SheetUtils.findColumnByName(sheet, metric.columnName);
                    if (metricColumn == null) {
                        continue;
                    }
                    const metricRange = sheet.getRange(row, metricColumn);
                    const foundIssues = metricsIssues.get().filter(metric.filter);
                    if (!foundIssues.length) {
                        metricRange.clearContent().setFontColor(null);
                        continue;
                    }
                    const metricIssueIds = foundIssues.map(issue => this.settings.issueIdGetter(issue));
                    const link = (_a = this.settings.issueIdsToUrl) === null || _a === void 0 ? void 0 : _a.call(null, metricIssueIds);
                    if (link != null) {
                        metricRange.setFormula(`=HYPERLINK("${link}", "${foundIssues.length}")`);
                    }
                    else {
                        metricRange.setFormula(`="${foundIssues.length}"`);
                    }
                    if (metric.color != null) {
                        metricRange.setFontColor(metric.color);
                    }
                }
            };
            calculateIssueMetrics(childIssues, this.settings.childIssueMetrics);
            calculateIssueMetrics(blockerIssues, this.settings.blockerIssueMetrics);
        }
        finally {
            issueIdRange.setBackground(null);
        }
    }
}
class Lazy {
    constructor(supplier) {
        this.supplier = supplier;
    }
    get() {
        if (this.supplier != null) {
            this.value = this.supplier();
            this.supplier = null;
        }
        return this.value;
    }
}
class RangeUtils {
    static doesRangeHaveColumn(range, columnName) {
        if (range == null) {
            return false;
        }
        const sheet = range.getSheet();
        const columnToFind = SheetUtils.findColumnByName(sheet, columnName);
        if (columnToFind == null) {
            return false;
        }
        for (const y of Utils.range(1, range.getHeight())) {
            let hasMerge = false;
            for (const x of Utils.range(1, range.getWidth())) {
                const cell = range.getCell(y, x);
                if (cell.isPartOfMerge()) {
                    hasMerge = true;
                }
                const col = cell.getColumn();
                if (col === columnToFind) {
                    return true;
                }
            }
            if (!hasMerge) {
                break;
            }
        }
        return false;
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
                title: link.title,
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
    static getMatrix(settingsSheet, settingsScope) {
        if (Utils.isString(settingsSheet)) {
            settingsSheet = SheetUtils.getSheetByName(settingsSheet);
        }
        settingsScope = Utils.normalizeName(settingsScope);
        return ExecutionCache.getOrComputeCache(['settings', 'map', settingsSheet, settingsScope], () => {
            const scopeRow = this.findScopeRow(settingsSheet, settingsScope);
            const columns = [];
            const columnsValues = settingsSheet.getRange(scopeRow + 1, 1, settingsSheet.getLastColumn(), 1).getValues()[0];
            for (const column of columnsValues) {
                const name = column.toString().trim();
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
            for (const row of Utils.range(scopeRow + 2, settingsSheet.getLastRow())) {
                const item = new Map();
                const values = settingsSheet.getRange(row, 1, 1, columns.length).getValues()[0];
                for (let i = 0; i < columns.length; ++i) {
                    item[columns[i]] = values[i].toString().trim();
                }
                result.push(item);
            }
            return result;
        });
    }
    static getMap(settingsSheet, settingsScope) {
        if (Utils.isString(settingsSheet)) {
            settingsSheet = SheetUtils.getSheetByName(settingsSheet);
        }
        settingsScope = Utils.normalizeName(settingsScope);
        return ExecutionCache.getOrComputeCache(['settings', 'map', settingsSheet, settingsScope], () => {
            const scopeRow = this.findScopeRow(settingsSheet, settingsScope);
            const result = new Map();
            for (const row of Utils.range(scopeRow + 1, settingsSheet.getLastRow())) {
                const values = settingsSheet.getRange(row, 1, 1, 2).getValues()[0];
                const key = values[0].toString().trim();
                const value = values[1].toString().trim();
                if (!key.length) {
                    break;
                }
                result[key] = value;
            }
            return result;
        });
    }
    static findScopeRow(sheet, scope) {
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
class SheetUtils {
    static findSheetByName(sheetName) {
        if (!(sheetName === null || sheetName === void 0 ? void 0 : sheetName.length)) {
            return null;
        }
        sheetName = Utils.normalizeName(sheetName);
        return ExecutionCache.getOrComputeCache(['findSheetByName', sheetName], () => {
            for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
                const name = Utils.normalizeName(sheet.getSheetName());
                if (name === sheetName) {
                    return sheet;
                }
            }
            return null;
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
            return null;
        }
        if (Utils.isString(sheet)) {
            sheet = this.findSheetByName(sheet);
        }
        if (sheet == null) {
            return null;
        }
        columnName = Utils.normalizeName(columnName);
        return ExecutionCache.getOrComputeCache([sheet, columnName], () => {
            for (const col of Utils.range(1, sheet.getLastColumn())) {
                const name = Utils.normalizeName(sheet.getRange(1, col).getValue());
                if (name === columnName) {
                    return col;
                }
            }
            return null;
        });
    }
    static getColumnByName(sheet, columnName) {
        var _a;
        if (Utils.isString(sheet)) {
            sheet = this.getSheetByName(sheet);
        }
        return (_a = this.findColumnByName(sheet, columnName)) !== null && _a !== void 0 ? _a : (() => {
            throw new Error(`"${columnName}" can't be found on "${sheet.getSheetName()}" sheet`);
        })();
    }
}
class Utils {
    static entryPoint(action) {
        try {
            return action();
        }
        catch (e) {
            console.error(e);
            throw e;
        }
    }
    static *range(startIncluding, endIncluding) {
        for (let n = startIncluding; n <= endIncluding; ++n) {
            yield n;
        }
    }
    static normalizeName(name) {
        return name.toString().trim().replaceAll(/\s+/g, ' ').toLowerCase();
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
    static isString(value) {
        return typeof value === 'string';
    }
    static isFunction(value) {
        return typeof value === 'function';
    }
}
