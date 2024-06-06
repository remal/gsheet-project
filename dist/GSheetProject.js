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
class GSheetProject {
    static reloadIssues() {
        EntryPoint.entryPoint(() => {
        });
    }
    static migrate() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrate();
        });
    }
    static refreshEverything() {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateIfNeeded();
            const sheet = SheetUtils.getSheetByName(GSheetProjectSettings.sheetName);
            const range = sheet.getRange(GSheetProjectSettings.firstDataRow, 1, Math.max(sheet.getLastRow() - GSheetProjectSettings.firstDataRow + 1, 1), sheet.getLastColumn());
            this._onEditRange(range);
        });
    }
    static cleanup() {
        EntryPoint.entryPoint(() => {
            ProtectionLocks.releaseExpiredLocks();
        });
    }
    static onOpen(event) {
        EntryPoint.entryPoint(() => {
            SheetLayouts.migrateIfNeeded();
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
        });
    }
    static _onRemoveColumn() {
        this.migrate();
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
            DoneLogic.executeDoneLogic(range);
            DefaultFormulas.insertDefaultFormulas(range);
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
GSheetProjectSettings.restoreUndoneEnd = false;
GSheetProjectSettings.sheetName = "Projects";
GSheetProjectSettings.iconColumnName = "icon";
GSheetProjectSettings.doneColumnName = "Done";
GSheetProjectSettings.milestoneColumnName = "Milestone";
GSheetProjectSettings.typeColumnName = "Type";
GSheetProjectSettings.issueColumnName = "Issue";
GSheetProjectSettings.issuesRangeName = "Issues";
GSheetProjectSettings.childIssueColumnName = "Child\nIssue";
GSheetProjectSettings.childIssuesRangeName = "ChildIssues";
GSheetProjectSettings.titleColumnName = "Title";
GSheetProjectSettings.teamColumnName = "Team";
GSheetProjectSettings.estimateColumnName = "Estimate\n(days)";
GSheetProjectSettings.deadlineColumnName = "Deadline";
GSheetProjectSettings.startColumnName = "Start";
GSheetProjectSettings.endColumnName = "End";
//static issueHashColumnName: string = "Issue Hash"
GSheetProjectSettings.settingsSheetName = "Settings";
GSheetProjectSettings.indent = 4;
GSheetProjectSettings.taskTrackers = [];
class AbstractIssueLogic {
    static _processRange(range) {
        if (![GSheetProjectSettings.issueColumnName, GSheetProjectSettings.titleColumnName].some(columnName => RangeUtils.doesRangeHaveSheetColumn(range, GSheetProjectSettings.sheetName, columnName))) {
            return null;
        }
        const sheet = range.getSheet();
        ProtectionLocks.lockAllColumns(sheet);
        range = RangeUtils.withMinMaxRows(range, GSheetProjectSettings.firstDataRow, sheet.getLastRow());
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
        return SheetUtils.getColumnsStringValues(sheet, {
            issues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName),
            childIssues: SheetUtils.getColumnByName(sheet, GSheetProjectSettings.childIssueColumnName),
        }, startRow, endRow);
    }
    static _getStringValues(range, column) {
        return RangeUtils.toColumnRange(range, column).getValues()
            .map(it => it[0].toString());
    }
    static _getFormulas(range, column) {
        return RangeUtils.toColumnRange(range, column).getFormulas()
            .map(it => it[0]);
    }
}
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
        const generateAndInsertFormulas = (column, formulaGenerator) => {
            const values = this._getStringValues(range, column);
            const formulas = this._getFormulas(range, column);
            for (let row = startRow; row <= endRow; ++row) {
                const index = row - startRow;
                if (!issues[index].length && !childIssues[index].length) {
                    if (formulas[index].length) {
                        sheet.getRange(row, column).setFormula('');
                    }
                    continue;
                }
                if (!values[index].length && !formulas[index].length) {
                    const formula = Utils.processFormula(formulaGenerator(row));
                    sheet.getRange(row, column).setFormula(formula);
                }
            }
        };
        const estimateColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.estimateColumnName);
        const startColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.startColumnName);
        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName);
        generateAndInsertFormulas(endColumn, row => {
            const startA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, startColumn));
            const estimateA1Notation = RangeUtils.getAbsoluteA1Notation(sheet.getRange(row, estimateColumn));
            return `
                =IF(
                    OR(
                        ISBLANK(${startA1Notation}),
                        ISBLANK(${estimateA1Notation})
                    ),
                    "",
                    WORKDAY(${startA1Notation}, ${estimateA1Notation})
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
class DoneLogic extends AbstractIssueLogic {
    static executeDoneLogic(range) {
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
        const hasIssue = (row) => {
            var _a, _b;
            const index = row - startRow;
            return !!((_a = issues[index]) === null || _a === void 0 ? void 0 : _a.length) || !!((_b = childIssues[index]) === null || _b === void 0 ? void 0 : _b.length);
        };
        const doneColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.doneColumnName);
        let doneValues = this._getStringValues(range, doneColumn);
        const checkboxesA1Notations = Array.from(Utils.range(startRow, endRow))
            .filter(row => hasIssue(row))
            .map(row => sheet.getRange(row, doneColumn).getA1Notation());
        if (checkboxesA1Notations.length) {
            sheet.getRangeList(checkboxesA1Notations).insertCheckboxes();
        }
        const notCheckboxesA1Notations = Array.from(Utils.range(startRow, endRow))
            .filter(row => !hasIssue(row))
            .filter(row => { var _a; return (_a = doneValues[row - startRow]) === null || _a === void 0 ? void 0 : _a.length; })
            .map(row => sheet.getRange(row, doneColumn).getA1Notation());
        if (notCheckboxesA1Notations.length) {
            sheet.getRangeList(notCheckboxesA1Notations).removeCheckboxes().setValue('');
        }
        doneValues = this._getStringValues(range, doneColumn);
        const endColumn = SheetUtils.getColumnByName(sheet, GSheetProjectSettings.endColumnName);
        for (let row = startRow; row <= endRow; ++row) {
            if (!hasIssue(row)) {
                continue;
            }
            const index = row - startRow;
            let doneValue = doneValues[index].toLowerCase();
            const endRange = sheet.getRange(row, endColumn);
            const rowRange = sheet.getRange(`${row}:${row}`);
            if (doneValue === 'true') {
                const endValue = endRange.getValue();
                let endDate;
                if (Utils.isString(endValue)) {
                    endDate = new Date(Number.isNaN(endValue) ? endValue : parseFloat(endValue));
                }
                else if (Utils.isNumber(endValue)) {
                    endDate = new Date(endValue);
                }
                else {
                    try {
                        endDate = new Date(endValue.toString());
                    }
                    catch (e) {
                        console.warn(`Can't get date from ${endRange.getA1Notation()}`);
                        continue;
                    }
                }
                if (GSheetProjectSettings.restoreUndoneEnd) {
                    const developerMetadata = rowRange.getDeveloperMetadata();
                    const previousFormulaMetadata = developerMetadata.find(it => it.getKey() === `${DoneLogic.name}|before-done-end-formula`);
                    if (previousFormulaMetadata != null) {
                        rowRange.addDeveloperMetadata(`${DoneLogic.name}|before-done-end-formula`, endRange.getFormula());
                    }
                    const previousValueMetadata = developerMetadata.find(it => it.getKey() === `${DoneLogic.name}|before-done-end-value`);
                    if (previousValueMetadata != null) {
                        rowRange.addDeveloperMetadata(`${DoneLogic.name}|before-done-end-value`, endDate.toString());
                    }
                }
                const now = new Date();
                if (now.getTime() < endDate.getTime()) {
                    endRange.setValue(now);
                }
                else {
                    endRange.setValue(endDate);
                }
            }
            else if (GSheetProjectSettings.restoreUndoneEnd) {
                const developerMetadata = rowRange.getDeveloperMetadata();
                const previousFormulaMetadata = developerMetadata.find(it => it.getKey() === `${DoneLogic.name}|before-done-end-formula`);
                const previousValueMetadata = developerMetadata.find(it => it.getKey() === `${DoneLogic.name}|before-done-end-value`);
                try {
                    const previousFormula = previousFormulaMetadata === null || previousFormulaMetadata === void 0 ? void 0 : previousFormulaMetadata.getValue();
                    const previousValue = previousValueMetadata === null || previousValueMetadata === void 0 ? void 0 : previousValueMetadata.getValue();
                    if (previousFormulaMetadata != null && (previousFormula === null || previousFormula === void 0 ? void 0 : previousFormula.length)) {
                        endRange.setFormula(previousFormula);
                    }
                    else if (previousValueMetadata != null) {
                        if (previousValue === null || previousValue === void 0 ? void 0 : previousValue.length) {
                            endRange.setValue(new Date(previousValue));
                        }
                        else {
                            endRange.setValue('');
                        }
                    }
                }
                finally {
                    previousFormulaMetadata === null || previousFormulaMetadata === void 0 ? void 0 : previousFormulaMetadata.remove();
                    previousValueMetadata === null || previousValueMetadata === void 0 ? void 0 : previousValueMetadata.remove();
                }
            }
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
class IssueHierarchyFormatter {
    static formatHierarchy(range) {
        if (![GSheetProjectSettings.childIssueColumnName].some(columnName => RangeUtils.doesRangeHaveSheetColumn(range, GSheetProjectSettings.sheetName, columnName))) {
            return;
        }
        let issuesRange = RangeUtils.toColumnRange(range, GSheetProjectSettings.issueColumnName);
        if (issuesRange != null) {
            issuesRange = RangeUtils.withMinRow(issuesRange, GSheetProjectSettings.firstDataRow);
            const issues = issuesRange.getValues()
                .map(it => { var _a; return (_a = it[0]) === null || _a === void 0 ? void 0 : _a.toString(); })
                .filter(it => it === null || it === void 0 ? void 0 : it.length)
                .filter(Utils.distinct());
            if (issues.length) {
                this.reorderIssuesAccordingToHierarchy(issues);
                this.formatHierarchyIssues(issues);
            }
        }
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
        const notEmptyIssues = issues.filter(it => it === null || it === void 0 ? void 0 : it.length).map(it => it);
        const notEmptyUniqueIssues = notEmptyIssues.filter(Utils.distinct());
        if (notEmptyIssues.length === notEmptyUniqueIssues.length) {
            return;
        }
        Utils.trimArrayEndBy(issues, it => !(it === null || it === void 0 ? void 0 : it.length));
        childIssues.length = issues.length;
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
        const notEmptyIssues = issues.filter(it => it === null || it === void 0 ? void 0 : it.length).map(it => it);
        const notEmptyUniqueIssues = notEmptyIssues.filter(Utils.distinct());
        if (notEmptyIssues.length === notEmptyUniqueIssues.length) {
            return;
        }
        const { milestoneFormulas, typeFormulas, deadlineFormulas, } = SheetUtils.getColumnsFormulas(sheet, {
            milestoneFormulas: milestonesColumn,
            typeFormulas: typesColumn,
            deadlineFormulas: deadlinesColumn,
        }, GSheetProjectSettings.firstDataRow);
        Utils.trimArrayEndBy(issues, it => !(it === null || it === void 0 ? void 0 : it.length));
        childIssues.length = issues.length;
        milestones.length = issues.length;
        types.length = issues.length;
        deadlines.length = issues.length;
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
                    const notations = indexes.map(index => {
                        const row = GSheetProjectSettings.firstDataRow + index;
                        return sheet.getRange(row, titlesColumn).getA1Notation();
                    });
                    const numberFormat = indent > 0
                        ? ' '.repeat(indent) + '@'
                        : '@';
                    sheet.getRangeList(notations).setNumberFormat(numberFormat);
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
    static lockAllColumns(sheet) {
        const sheetId = sheet.getSheetId();
        if (this._allColumnsProtections.has(sheetId)) {
            return;
        }
        const range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
        const protection = range.protect()
            .setDescription(`lock|columns|all|${new Date().getTime()}`)
            .setWarningOnly(true);
        this._allColumnsProtections.set(sheetId, protection);
    }
    static lockAllRows(sheet) {
        const sheetId = sheet.getSheetId();
        if (this._allRowsProtections.has(sheetId)) {
            return;
        }
        const range = sheet.getRange(1, sheet.getMaxColumns(), sheet.getMaxRows(), 1);
        const protection = range.protect()
            .setDescription(`lock|rows|all|${new Date().getTime()}`)
            .setWarningOnly(true);
        this._allRowsProtections.set(sheetId, protection);
    }
    static lockRows(sheet, rowsToLock) {
        if (rowsToLock <= 0) {
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
        if (maxLockedRow < rowsToLock) {
            const range = sheet.getRange(1, sheet.getMaxColumns(), rowsToLock, 1);
            const protection = range.protect()
                .setDescription(`lock|rows|${rowsToLock}|${new Date().getTime()}`)
                .setWarningOnly(true);
            rowsProtections.set(rowsToLock, protection);
        }
    }
    static release() {
        this._allColumnsProtections.forEach(protection => protection.remove());
        this._allColumnsProtections.clear();
        this._allRowsProtections.forEach(protection => protection.remove());
        this._allRowsProtections.clear();
        this._rowsProtections.forEach(protections => Array.from(protections.values()).forEach(protection => protection.remove()));
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
            .replace(/[A-Z]+/, '$$$&')
            .replace(/\d+/, '$$$&');
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
        const rowDiff = minRow - range.getRow();
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
        return `${((_a = this.constructor) === null || _a === void 0 ? void 0 : _a.name) || Utils.normalizeName(this.sheetName)}:migrateColumns:`;
    }
    get _documentFlag() {
        return `${this._documentFlagPrefix}5a18403f4a9f94173ce373e3a757159e986fb9daa46561e733ae29551c9cd7cf:${GSheetProjectSettings.computeStringSettingsHash()}`;
    }
    migrateIfNeeded() {
        if (DocumentFlags.isSet(this._documentFlag)) {
            return;
        }
        this.migrate();
    }
    migrate() {
        var _a, _b, _c, _d, _e;
        const sheet = this.sheet;
        const columns = this.columns.reduce((map, info) => {
            map.set(Utils.normalizeName(info.name), info);
            return map;
        }, new Map());
        if (!columns.size) {
            return;
        }
        ProtectionLocks.lockAllColumns(sheet);
        let lastColumn = sheet.getLastColumn();
        const maxRows = sheet.getMaxRows();
        const existingNormalizedNames = sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn)
            .getValues()[0]
            .map(it => it === null || it === void 0 ? void 0 : it.toString())
            .map(it => (it === null || it === void 0 ? void 0 : it.length) ? Utils.normalizeName(it) : '');
        for (const [columnName, info] of columns.entries()) {
            if (existingNormalizedNames.includes(columnName)) {
                continue;
            }
            console.info(`Adding "${info.name}" column`);
            ++lastColumn;
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
            if (info.defaultFormat != null) {
                sheet.getRange(GSheetProjectSettings.firstDataRow, lastColumn, maxRows, 1)
                    .setNumberFormat(info.defaultFormat);
            }
            if ((_a = info.defaultHorizontalAlignment) === null || _a === void 0 ? void 0 : _a.length) {
                sheet.getRange(GSheetProjectSettings.firstDataRow, lastColumn, maxRows, 1)
                    .setHorizontalAlignment(info.defaultHorizontalAlignment);
            }
            if (info.hiddenByDefault) {
                sheet.hideColumns(lastColumn);
            }
            existingNormalizedNames.push(columnName);
        }
        const existingFormulas = new Lazy(() => sheet.getRange(GSheetProjectSettings.titleRow, 1, 1, lastColumn).getFormulas()[0]);
        for (const [columnName, info] of columns.entries()) {
            const index = existingNormalizedNames.indexOf(columnName);
            if (index < 0) {
                continue;
            }
            const column = index + 1;
            if ((_b = info.arrayFormula) === null || _b === void 0 ? void 0 : _b.length) {
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
            if ((_c = info.rangeName) === null || _c === void 0 ? void 0 : _c.length) {
                SpreadsheetApp.getActiveSpreadsheet().setNamedRange(info.rangeName, range);
            }
            let dataValidation = (_e = (_d = info.dataValidation) === null || _d === void 0 ? void 0 : _d.call(info)) !== null && _e !== void 0 ? _e : null;
            if (dataValidation != null) {
                if (dataValidation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CUSTOM_FORMULA) {
                    const formula = Utils.processFormula(dataValidation.getCriteriaValues()[0].toString());
                    dataValidation = dataValidation.copy()
                        .requireFormulaSatisfied(formula)
                        .build();
                }
            }
            range.setDataValidation(dataValidation);
        }
        sheet.getRange(1, 1, lastColumn, 1)
            .setHorizontalAlignment('center')
            .setFontWeight('bold')
            .setNumberFormat('');
        DocumentFlags.set(this._documentFlag);
        DocumentFlags.cleanupByPrefix(this._documentFlagPrefix);
        const waitForAllDataExecutionsCompletion = SpreadsheetApp.getActiveSpreadsheet()['waitForAllDataExecutionsCompletion'];
        if (Utils.isFunction(waitForAllDataExecutionsCompletion)) {
            try {
                waitForAllDataExecutionsCompletion(10);
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
                name: GSheetProjectSettings.doneColumnName,
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
                            NOT(ISBLANK(${GSheetProjectSettings.childIssuesRangeName})),
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
                name: GSheetProjectSettings.teamColumnName,
                defaultFormat: '',
                defaultHorizontalAlignment: 'left',
            },
            {
                name: GSheetProjectSettings.estimateColumnName,
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
            },
            {
                name: GSheetProjectSettings.deadlineColumnName,
                defaultFormat: 'yyyy-MM-dd',
                defaultHorizontalAlignment: 'center',
            },
            /*
            {
                name: GSheetProjectSettings.projectsIssueHashColumnName,
                hiddenByDefault: true,
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
    static get instances() {
        return [
            SheetLayoutProjects.instance,
            SheetLayoutSettings.instance,
        ];
    }
    static migrateIfNeeded() {
        this.instances.forEach(instance => instance.migrateIfNeeded());
    }
    static migrate() {
        this.instances.forEach(instance => instance.migrate());
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
    static isGridSheet(sheet) {
        if (Utils.isString(sheet)) {
            sheet = this.findSheetByName(sheet);
        }
        if (sheet == null) {
            return false;
        }
        return sheet.getType() === SpreadsheetApp.SheetType.GRID;
    }
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
        ProtectionLocks.lockAllColumns(sheet);
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
        const lastRow = sheet.getLastRow();
        if (minRow > lastRow) {
            return sheet.getRange(minRow, column);
        }
        const rows = lastRow - minRow + 1;
        return sheet.getRange(minRow, column, rows, 1);
    }
    static getColumnsValues(sheet, columns, minRow, maxRow) {
        const getter = range => range.getValues();
        return this._getColumnsProps(sheet, columns, getter, minRow, maxRow);
    }
    static getColumnsStringValues(sheet, columns, minRow, maxRow) {
        const getter = range => range.getValues();
        const result = this._getColumnsProps(sheet, columns, getter, minRow, maxRow);
        for (const [key, values] of Object.entries(result)) {
            result[key] = values.map(value => value.toString());
        }
        return result;
    }
    static getColumnsFormulas(sheet, columns, minRow, maxRow) {
        const getter = range => range.getFormulas();
        return this._getColumnsProps(sheet, columns, getter, minRow, maxRow);
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
            maxRow = sheet.getLastRow();
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
        const lastColumn = sheet.getLastColumn();
        if (minColumn > lastColumn) {
            return sheet.getRange(row, minColumn);
        }
        const columns = lastColumn - minColumn + 1;
        return sheet.getRange(row, minColumn, 1, columns);
    }
}
class TaskTracker {
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
            .replaceAll(/#SELF\b/g, 'INDIRECT(ADDRESS(ROW(), COLUMN()))')
            .split(/[\r\n]+/)
            .map(line => line.trim())
            .filter(line => line.length)
            .map(line => line + (line.endsWith(',') || line.endsWith(';') ? ' ' : ''))
            .join('');
    }
    static escapeFormulaString(string) {
        return string.replaceAll(/"/g, '""');
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
    static numericAsc() {
        return (n1, n2) => n1 - n2;
    }
    static numericDesc() {
        return (n1, n2) => n2 - n1;
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