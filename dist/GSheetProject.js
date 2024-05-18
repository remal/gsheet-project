class GSheetProject {
    static reloadIssues() {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache();
            IssueLoader.loadAllIssues();
        });
    }
    static onOpen(event) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache();
        });
    }
    static onChange(event) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache();
            HierarchyFormatter.formatAllHierarchy();
        });
    }
    static onEdit(event) {
        this.onEditRange(event === null || event === void 0 ? void 0 : event.range);
    }
    static onFormSubmit(event) {
        this.onEditRange(event === null || event === void 0 ? void 0 : event.range);
    }
    static onEditRange(range) {
        Utils.entryPoint(() => {
            ExecutionCache.resetCache();
            if (range != null) {
                IssueIdFormatter.formatIssueId(range);
                HierarchyFormatter.formatHierarchy(range);
                IssueLoader.loadIssues(range);
            }
        });
    }
}
class GSheetProjectSettings {
}
GSheetProjectSettings.firstDataRow = 2;
GSheetProjectSettings.settingsSheetName = "Settings";
GSheetProjectSettings.issueIdColumnName = "Issue";
GSheetProjectSettings.parentIssueIdColumnName = "Parent Issue";
GSheetProjectSettings.isDoneColumnName = "Done";
GSheetProjectSettings.issueIdsExtractor = () => Utils.throwNotConfigured('issueIdsExtractor');
GSheetProjectSettings.issueIdDecorator = () => Utils.throwNotConfigured('issueIdDecorator');
GSheetProjectSettings.issueIdToUrl = () => Utils.throwNotConfigured('issueIdToUrl');
GSheetProjectSettings.issueIdsToUrl = () => Utils.throwNotConfigured('issueIdsToUrl');
GSheetProjectSettings.issuesLoader = () => Utils.throwNotConfigured('issuesLoader');
GSheetProjectSettings.childIssuesLoader = () => Utils.throwNotConfigured('childIssuesLoader');
GSheetProjectSettings.blockerIssuesLoader = () => Utils.throwNotConfigured('blockerIssuesLoader');
GSheetProjectSettings.issueIdGetter = () => Utils.throwNotConfigured('issueIdGetter');
GSheetProjectSettings.idDoneCalculator = () => Utils.throwNotConfigured('idDoneCalculator');
GSheetProjectSettings.stringFields = {};
GSheetProjectSettings.booleanFields = {};
GSheetProjectSettings.childIssueMetrics = [];
GSheetProjectSettings.blockerIssueMetrics = [];
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
            return this.data.get(stringKey);
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
class HierarchyFormatter {
    static formatHierarchy(range) {
        if (!RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.issueIdColumnName)
            && !RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.parentIssueIdColumnName)) {
            return;
        }
        this.formatSheetHierarchy(range.getSheet());
    }
    static formatAllHierarchy() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            this.formatSheetHierarchy(sheet);
        }
    }
    static formatSheetHierarchy(sheet) {
        const issueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.issueIdColumnName);
        const parentIssueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.parentIssueIdColumnName);
        if (issueIdColumn == null || parentIssueIdColumn == null) {
            return;
        }
        const lastRow = sheet.getLastRow();
        const getAllIds = (column) => {
            return sheet.getRange(GSheetProjectSettings.firstDataRow, column, lastRow - GSheetProjectSettings.firstDataRow, 1)
                .getValues()
                .map(cols => cols[0].toString())
                .map(text => GSheetProjectSettings.issueIdsExtractor(text));
        };
        // group children:
        while (true) {
            const allParentIssueIds = getAllIds(parentIssueIdColumn);
            let isChanged = false;
            for (let index = allParentIssueIds.length - 1; 0 <= index; --index) {
                const parentIssueIds = allParentIssueIds[index];
                if (!(parentIssueIds === null || parentIssueIds === void 0 ? void 0 : parentIssueIds.length)) {
                    continue;
                }
                let previousIndex = null;
                for (let prevIndex = index - 1; 0 <= prevIndex; --prevIndex) {
                    const prevParentIssueIds = allParentIssueIds[prevIndex];
                    if (Utils.arrayEquals(parentIssueIds, prevParentIssueIds)) {
                        previousIndex = prevIndex;
                        break;
                    }
                }
                if (previousIndex != null && previousIndex < index - 1) {
                    const newIndex = previousIndex + 1;
                    const row = GSheetProjectSettings.firstDataRow + index;
                    const newRow = GSheetProjectSettings.firstDataRow + newIndex;
                    sheet.moveRows(sheet.getRange(row, 1), newRow);
                    isChanged = true;
                }
            }
            if (!isChanged) {
                break;
            }
        }
        // move children:
        while (true) {
            const allIssueIds = getAllIds(issueIdColumn);
            const allParentIssueIds = getAllIds(parentIssueIdColumn);
            let isChanged = false;
            for (let index = 0; index < allParentIssueIds.length; ++index) {
                const currentIndex = index;
                const parentIssueIds = allParentIssueIds[currentIndex];
                if (!(parentIssueIds === null || parentIssueIds === void 0 ? void 0 : parentIssueIds.length)) {
                    continue;
                }
                let groupSize = 1;
                for (index; index < allParentIssueIds.length - 1; ++index) {
                    const nextParentIssueIds = allParentIssueIds[index + 1];
                    if (Utils.arrayEquals(parentIssueIds, nextParentIssueIds)) {
                        ++groupSize;
                    }
                    else {
                        break;
                    }
                }
                const issueIndex = allIssueIds.findIndex(ids => ids === null || ids === void 0 ? void 0 : ids.some(id => parentIssueIds.includes(id)));
                if (issueIndex < 0 || issueIndex == currentIndex || issueIndex == currentIndex - 1) {
                    continue;
                }
                const newIndex = issueIndex + 1;
                const row = GSheetProjectSettings.firstDataRow + currentIndex;
                const newRow = GSheetProjectSettings.firstDataRow + newIndex;
                sheet.moveRows(sheet.getRange(row, 1, groupSize, 1), newRow);
                break;
            }
            if (!isChanged) {
                break;
            }
        }
    }
}
class IssueIdFormatter {
    static formatIssueId(range) {
        var _a;
        const columnNames = [
            GSheetProjectSettings.issueIdColumnName,
            GSheetProjectSettings.parentIssueIdColumnName,
        ];
        for (const y of Utils.range(1, range.getHeight())) {
            for (const x of Utils.range(1, range.getWidth())) {
                const cell = range.getCell(y, x);
                if (!columnNames.some(name => RangeUtils.doesRangeHaveColumn(cell, name))) {
                    continue;
                }
                const ids = (_a = GSheetProjectSettings.issueIdsExtractor(cell.getValue())) !== null && _a !== void 0 ? _a : [];
                const links = ids.map(id => {
                    return {
                        url: GSheetProjectSettings.issueIdToUrl(id),
                        title: GSheetProjectSettings.issueIdDecorator(id),
                    };
                });
                cell.setRichTextValue(RichTextUtils.createLinksValue(links));
            }
        }
    }
}
class IssueLoader {
    static loadIssues(range) {
        if (!RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.issueIdColumnName)) {
            return;
        }
        const sheet = range.getSheet();
        const rows = Array.from(Utils.range(1, range.getHeight()))
            .map(y => range.getCell(y, 1).getRow())
            .filter(row => row >= GSheetProjectSettings.firstDataRow)
            .filter(Utils.distinct);
        for (const row of rows) {
            this.loadIssuesForRow(sheet, row);
        }
    }
    static loadAllIssues() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            const hasIssueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.issueIdColumnName) != null;
            if (!hasIssueIdColumn) {
                continue;
            }
            for (const row of Utils.range(GSheetProjectSettings.firstDataRow, sheet.getLastRow())) {
                this.loadIssuesForRow(sheet, row);
            }
        }
    }
    static loadIssuesForRow(sheet, row) {
        if (row < GSheetProjectSettings.firstDataRow) {
            return;
        }
        const issueIdColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.issueIdColumnName);
        const issueIdRange = sheet.getRange(row, issueIdColumn);
        const issueIds = GSheetProjectSettings.issueIdsExtractor(issueIdRange.getValue());
        if (!(issueIds === null || issueIds === void 0 ? void 0 : issueIds.length)) {
            return;
        }
        console.log(`"${sheet.getSheetName()}" sheet: processing row #${row}`);
        issueIdRange.setBackground('#eee');
        try {
            const rootIssues = GSheetProjectSettings.issuesLoader(issueIds);
            const childIssues = new Lazy(() => {
                return GSheetProjectSettings.childIssuesLoader(issueIds)
                    .filter(issue => !issueIds.includes(GSheetProjectSettings.issueIdGetter(issue)));
            });
            const blockerIssues = new Lazy(() => {
                const ids = rootIssues.concat(childIssues.get())
                    .map(issue => GSheetProjectSettings.issueIdGetter(issue));
                return GSheetProjectSettings.blockerIssuesLoader(ids)
                    .filter(issue => !issueIds.includes(GSheetProjectSettings.issueIdGetter(issue)));
            });
            const isDoneColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.isDoneColumnName);
            if (isDoneColumn != null) {
                const isDone = GSheetProjectSettings.idDoneCalculator(rootIssues, childIssues.get());
                sheet.getRange(row, isDoneColumn).setValue(isDone ? 'Yes' : '');
            }
            for (const [columnName, getter] of Object.entries(GSheetProjectSettings.stringFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName);
                if (fieldColumn != null) {
                    sheet.getRange(row, fieldColumn).setValue(rootIssues
                        .map(getter)
                        .join('\n'));
                }
            }
            for (const [columnName, getter] of Object.entries(GSheetProjectSettings.booleanFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName);
                if (fieldColumn != null) {
                    const isTrue = rootIssues.every(getter);
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
                    const metricIssueIds = foundIssues.map(issue => GSheetProjectSettings.issueIdGetter(issue));
                    const link = (_a = GSheetProjectSettings.issueIdsToUrl) === null || _a === void 0 ? void 0 : _a.call(null, metricIssueIds);
                    if (link != null) {
                        metricRange.setFormula(`=HYPERLINK("${link}", "${foundIssues.length}")`);
                    }
                    else {
                        metricRange.setFormula(`="${foundIssues.length}"`);
                    }
                    metricRange.setFontColor(metric.color);
                }
            };
            calculateIssueMetrics(childIssues, GSheetProjectSettings.childIssueMetrics);
            calculateIssueMetrics(blockerIssues, GSheetProjectSettings.blockerIssueMetrics);
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
        return ExecutionCache.getOrComputeCache(['settings', 'matrix', settingsSheet, settingsScope], () => {
            const scopeRow = this.findScopeRow(settingsSheet, settingsScope);
            const columns = [];
            const columnsValues = settingsSheet.getRange(scopeRow + 1, 1, 1, settingsSheet.getLastColumn()).getValues()[0];
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
                    item.set(columns[i], values[i].toString().trim());
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
                result.set(key, value);
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
        return ExecutionCache.getOrComputeCache(['findColumnByName', sheet, columnName], () => {
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
            throw new Error(`"${sheet.getSheetName()}" sheet: "${columnName}" column can't be found`);
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
    static isString(value) {
        return typeof value === 'string';
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
