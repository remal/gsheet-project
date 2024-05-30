class GSheetProject {
    static reloadIssues() {
        EntryPoint.entryPoint(() => {
        });
    }
    static migrateColumns() {
        EntryPoint.entryPoint(() => {
            ProjectsSheetLayout.instance.migrateColumns();
        });
    }
    static onOpen(event) {
        EntryPoint.entryPoint(() => {
            ProjectsSheetLayout.instance.migrateColumns();
        });
    }
    static onChange(event) {
        var _a, _b;
        if (!['INSERT_ROW', 'OTHER'].includes((_b = (_a = event === null || event === void 0 ? void 0 : event.changeType) === null || _a === void 0 ? void 0 : _a.toString()) !== null && _b !== void 0 ? _b : '')) {
            return;
        }
        EntryPoint.entryPoint(() => {
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
        });
    }
}

class GSheetProjectSettings {
    static computeSettingsHash() {
        const hashableValues = {};
        for (const [key, value] of Object.entries(GSheetProjectSettings)) {
            if (value == null
                || typeof value === 'function'
                || typeof value === 'object') {
                continue;
            }
            hashableValues[key] = value;
        }
        const json = JSON.stringify(hashableValues);
        return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, json);
    }
}
GSheetProjectSettings.titleRow = 1;
GSheetProjectSettings.firstDataRow = 2;
GSheetProjectSettings.settingsSheetName = "Settings";
GSheetProjectSettings.projectsSheetName = "Projects";
GSheetProjectSettings.projectsIssueColumnName = "Issue";
GSheetProjectSettings.projectsIssuesRangeName = "Issues";
GSheetProjectSettings.projectsIssueHashColumnName = "Issue Hash";
GSheetProjectSettings.projectsIssueHashesRangeName = "IssueHashes";

class AbstractSheetLayout {
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
        const cacheKey = [
            ((_a = this.constructor) === null || _a === void 0 ? void 0 : _a.name) || Utils.normalizeName(this.sheetName),
            'migrateColumns',
            'ae27689a90b3b6c7ceab1f3a807dbe43f4ebf5cbe1c968c476d212d243382660',
            GSheetProjectSettings.computeSettingsHash(),
        ].join(':').replace(/^(.{1,250}).*$/, '$1');
        const cache = PropertiesService.getDocumentProperties();
        if (cache != null) {
            if (cache.getProperty(cacheKey) === 'true') {
                return;
            }
        }
        const sheet = this.sheet;
        ProtectionLocks.lockColumnsWithProtection(sheet);
        let lastColumn = sheet.getLastColumn();
        const maxRows = sheet.getMaxRows();
        const existingNormalizedNames = sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn)
            .getValues()[0]
            .map(it => it === null || it === void 0 ? void 0 : it.toString())
            .map(it => (it === null || it === void 0 ? void 0 : it.length) ? Utils.normalizeName(it) : '');
        for (const [columnName, info] of columns.entries()) {
            if (!existingNormalizedNames.includes(columnName)) {
                sheet.getRange(GSheetProjectSettings.titleRow, lastColumn)
                    .setValue(info.name);
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
                    .join('')
                    .trim();
                const formulaToExpect = `={"${Utils.escapeFormulaString(info.name)}", ${arrayFormulaNormalized}`;
                const formula = existingFormulas.get()[index];
                if (formula !== formulaToExpect) {
                    sheet.getRange(GSheetProjectSettings.titleRow, column)
                        .setFormula(formulaToExpect);
                }
            }
            if ((_c = info.rangeName) === null || _c === void 0 ? void 0 : _c.length) {
                SpreadsheetApp.getActiveSpreadsheet().setNamedRange(info.rangeName, sheet.getRange(GSheetProjectSettings.firstDataRow, column, maxRows - GSheetProjectSettings.firstDataRow, 1));
            }
        }
        if (cache != null) {
            cache.setProperty(cacheKey, 'true');
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
    return Utilities.base64EncodeWebSafe(digest);
}

class EntryPoint {
    static entryPoint(action) {
        try {
            ExecutionCache.resetCache();
            return action();
        }
        catch (e) {
            console.error(e);
            throw e;
        }
        finally {
            ProtectionLocks.release();
            ProtectionLocks.releaseExpiredLocks();
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

class ProjectsSheetLayout extends AbstractSheetLayout {
    get sheetName() {
        return GSheetProjectSettings.projectsSheetName;
    }
    get columns() {
        return [
            {
                name: GSheetProjectSettings.projectsIssueColumnName,
                rangeName: GSheetProjectSettings.projectsIssuesRangeName,
            },
            {
                name: GSheetProjectSettings.projectsIssueHashColumnName,
                arrayFormula: `
                    MAP(
                        ARRAYFORMULA(${GSheetProjectSettings.projectsIssuesRangeName}),
                        LAMBDA(issue, IF(ISBLANK(issue), "", ${SHA256.name}(issue)))
                    )
                `,
                rangeName: GSheetProjectSettings.projectsIssueHashesRangeName,
            },
        ];
    }
}
ProjectsSheetLayout.instance = new ProjectsSheetLayout();

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
