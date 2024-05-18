class GSheetProject {
    constructor(settings) {
        this.settings = settings;
    }
    onOpen(event) {
        ExecutionCache.resetCache();
    }
    onChange(event) {
        ExecutionCache.resetCache();
    }
    osEdit(event) {
        ExecutionCache.resetCache();
    }
}
class GSheetProjectSettings {
    constructor() {
        this.settingsSheetName = "Settings";
        this.issueColumnName = "Issue";
        this.parentIssueColumnName = "Parent Issue";
        this.issueIdsExtractor = (_) => {
            throw new Error('issueIdsExtractor is not set');
        };
        this.issueIdToLink = (_) => {
            throw new Error('issueIdToLink is not set');
        };
    }
}
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
    static isString(value) {
        return typeof value === 'string';
    }
    static isFunction(value) {
        return typeof value === 'function';
    }
}
