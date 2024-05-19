class GSheetProject {
    static reloadIssues() {
        Utils.entryPoint(() => {
            IssueLoader.loadAllIssues();
        });
    }
    static recalculateSchedule() {
        Utils.entryPoint(() => {
            Schedule.recalculateAllSchedules();
        });
    }
    static onOpen(event) {
        Utils.entryPoint(() => {
        });
    }
    static onChange(event) {
        Utils.entryPoint(() => {
            State.updateLastStructureChange();
            HierarchyFormatter.formatAllHierarchy();
            Schedule.recalculateAllSchedules();
        });
    }
    static onEdit(event) {
        this._onEditRange(event === null || event === void 0 ? void 0 : event.range);
    }
    static onFormSubmit(event) {
        this._onEditRange(event === null || event === void 0 ? void 0 : event.range);
    }
    static _onEditRange(range) {
        Utils.entryPoint(() => {
            if (range != null) {
                IssueIdFormatter.formatIssueId(range);
                HierarchyFormatter.formatHierarchy(range);
                IssueLoader.loadIssues(range);
                Schedule.recalculateSchedule(range);
            }
        });
    }
}
class GSheetProjectSettings {
}
GSheetProjectSettings.firstDataRow = 2;
GSheetProjectSettings.settingsSheetName = "Settings";
GSheetProjectSettings.settingsTeamsScope = "Teams";
GSheetProjectSettings.settingsScheduleScope = "Schedule";
GSheetProjectSettings.issueIdColumnName = "Issue";
GSheetProjectSettings.parentIssueIdColumnName = "Parent Issue";
GSheetProjectSettings.isDoneColumnName = "Done";
GSheetProjectSettings.estimateColumnName = "Estimate";
GSheetProjectSettings.laneColumnName = "Lane";
GSheetProjectSettings.startColumnName = "Start";
GSheetProjectSettings.endColumnName = "End";
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
GSheetProjectSettings.aggregatedBooleanFields = {};
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
class HierarchyFormatter {
    static formatHierarchy(range) {
        if (!RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.issueIdColumnName)
            && !RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.parentIssueIdColumnName)) {
            return;
        }
        this._formatSheetHierarchy(range.getSheet());
    }
    static formatAllHierarchy() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            this._formatSheetHierarchy(sheet);
        }
    }
    static _formatSheetHierarchy(sheet) {
        if (State.isStructureChanged())
            return;
        const issueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.issueIdColumnName);
        const parentIssueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.parentIssueIdColumnName);
        if (issueIdColumn == null || parentIssueIdColumn == null) {
            return;
        }
        const getAllIds = (column) => {
            return SheetUtils.getColumnRange(sheet, column, GSheetProjectSettings.firstDataRow)
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
                    if (State.isStructureChanged())
                        return;
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
                if (State.isStructureChanged())
                    return;
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
            this._loadIssuesForRow(sheet, row);
        }
    }
    static loadAllIssues() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            if (State.isStructureChanged())
                return;
            const hasIssueIdColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.issueIdColumnName) != null;
            if (!hasIssueIdColumn) {
                continue;
            }
            for (const row of Utils.range(GSheetProjectSettings.firstDataRow, sheet.getLastRow())) {
                this._loadIssuesForRow(sheet, row);
            }
        }
    }
    static _loadIssuesForRow(sheet, row) {
        if (row < GSheetProjectSettings.firstDataRow) {
            return;
        }
        if (State.isStructureChanged())
            return;
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
                if (State.isStructureChanged())
                    return;
                sheet.getRange(row, isDoneColumn).setValue(isDone ? 'Yes' : '');
            }
            for (const [columnName, getter] of Object.entries(GSheetProjectSettings.stringFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName);
                if (fieldColumn != null) {
                    if (State.isStructureChanged())
                        return;
                    sheet.getRange(row, fieldColumn).setValue(rootIssues
                        .map(getter)
                        .join('\n'));
                }
            }
            for (const [columnName, getter] of Object.entries(GSheetProjectSettings.booleanFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName);
                if (fieldColumn != null) {
                    const isTrue = rootIssues.every(getter);
                    if (State.isStructureChanged())
                        return;
                    sheet.getRange(row, fieldColumn).setValue(isTrue ? 'Yes' : '');
                }
            }
            for (const [columnName, getter] of Object.entries(GSheetProjectSettings.aggregatedBooleanFields)) {
                const fieldColumn = SheetUtils.findColumnByName(sheet, columnName);
                if (fieldColumn != null) {
                    const isTrue = getter(rootIssues, childIssues.get());
                    if (State.isStructureChanged())
                        return;
                    sheet.getRange(row, fieldColumn).setValue(isTrue ? 'Yes' : '');
                }
            }
            const calculateIssueMetrics = (metricsIssues, metrics) => {
                var _a, _b;
                for (const metric of metrics) {
                    const metricColumn = SheetUtils.findColumnByName(sheet, metric.columnName);
                    if (metricColumn == null) {
                        continue;
                    }
                    const metricRange = sheet.getRange(row, metricColumn);
                    const foundIssues = metricsIssues.get().filter(metric.filter);
                    if (State.isStructureChanged())
                        return;
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
                    metricRange.setFontColor((_b = metric.color) !== null && _b !== void 0 ? _b : null);
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
class Lane {
    constructor() {
        this._amounts = [];
        this._objects = [];
    }
    add(amount, object) {
        this._amounts.push(amount);
        this._objects.push(object);
        return this;
    }
    get sum() {
        return this._amounts.reduce((sum, amount) => sum + amount, 0);
    }
    *[Symbol.iterator]() {
        for (let i = 0; i < this._amounts.length; ++i) {
            yield [this._amounts[i], this._objects[i]];
        }
    }
    *amounts() {
        for (const item of this._amounts) {
            yield item;
        }
    }
    *objects() {
        for (const item of this._objects) {
            yield item;
        }
    }
}
class Lanes {
    constructor(lanesNumber) {
        this.lanes = [];
        for (let i = 0; i < lanesNumber; ++i) {
            this.lanes.push(new Lane());
        }
    }
    add(amount, object, callback) {
        if (!this.lanes.length) {
            return this;
        }
        let laneIndex = 0;
        let laneToAdd = this.lanes[laneIndex];
        let currentSum = laneToAdd.sum;
        for (let i = 1; i < this.lanes.length; ++i) {
            const lane = this.lanes[i];
            const sum = lane.sum;
            if (sum < currentSum) {
                laneIndex = i;
                laneToAdd = lane;
                currentSum = sum;
            }
        }
        laneToAdd.add(amount, object);
        if (callback != null) {
            callback(laneIndex, laneToAdd, amount, object);
        }
        return this;
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
class Schedule {
    static recalculateSchedule(range) {
        if (!RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.estimateColumnName)
            || !RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.startColumnName)
            || !RangeUtils.doesRangeHaveColumn(range, GSheetProjectSettings.endColumnName)) {
            return;
        }
        this._recalculateSheetSchedule(range.getSheet());
    }
    static recalculateAllSchedules() {
        for (const sheet of SpreadsheetApp.getActiveSpreadsheet().getSheets()) {
            this._recalculateSheetSchedule(sheet);
        }
    }
    static _recalculateSheetSchedule(sheet) {
        if (State.isStructureChanged())
            return;
        const estimateColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.estimateColumnName);
        if (estimateColumn == null) {
            return;
        }
        let lanesRange = null;
        const laneColumn = SheetUtils.findColumnByName(sheet, GSheetProjectSettings.laneColumnName);
        if (laneColumn != null) {
            lanesRange = SheetUtils.getColumnRange(sheet, laneColumn, GSheetProjectSettings.firstDataRow).setValue('');
        }
        const startsRange = SheetUtils.getColumnRange(sheet, GSheetProjectSettings.startColumnName, GSheetProjectSettings.firstDataRow).setValue('');
        const endsRange = SheetUtils.getColumnRange(sheet, GSheetProjectSettings.endColumnName, GSheetProjectSettings.firstDataRow).setValue('');
        const generalEstimatesRange = SheetUtils.getColumnRange(sheet, estimateColumn, GSheetProjectSettings.firstDataRow);
        const generalEstimates = generalEstimatesRange.getValues()
            .map(cols => cols[0].toString().trim());
        const allDaysEstimates = [];
        const allTeamDaysEstimates = new Map;
        const invalidEstimateRows = [];
        generalEstimates.forEach((generalEstimate, index) => {
            var _a, _b;
            if (!generalEstimate.length) {
                return;
            }
            let isTeamFound = false;
            for (const team of Team.getAll()) {
                const regex = new RegExp(`^${Utils.escapeRegex(team.id)}\\s*:\\s*(\\d+)\\s*([dw])?$`, 'i');
                const match = generalEstimate.match(regex);
                if (match == null) {
                    continue;
                }
                const amount = parseInt(match[1]);
                const unit = (_b = (_a = match[2]) === null || _a === void 0 ? void 0 : _a.toUpperCase()) !== null && _b !== void 0 ? _b : 'D';
                let days;
                if (unit === 'W') {
                    days = amount * 5;
                }
                else {
                    days = amount;
                }
                const row = GSheetProjectSettings.firstDataRow + index;
                const canonicalGeneralEstimate = `${team.id}: ${amount}${unit !== 'D' ? unit : ''}`;
                if (canonicalGeneralEstimate !== generalEstimate) {
                    if (State.isStructureChanged())
                        return;
                    sheet.getRange(row, estimateColumn).setValue(canonicalGeneralEstimate);
                }
                let teamDayEstimates = allTeamDaysEstimates.get(team.id);
                if (teamDayEstimates == null) {
                    teamDayEstimates = [];
                    allTeamDaysEstimates.set(team.id, teamDayEstimates);
                }
                const dayEstimate = {
                    index: index,
                    row: row,
                    daysEstimate: days,
                    teamId: team.id,
                };
                allDaysEstimates.push(dayEstimate);
                teamDayEstimates.push(dayEstimate);
                isTeamFound = true;
                break;
            }
            if (!isTeamFound) {
                invalidEstimateRows.push(GSheetProjectSettings.firstDataRow + index);
            }
        });
        if (State.isStructureChanged())
            return;
        generalEstimatesRange.setBackground(null);
        const invalidEstimateNotations = invalidEstimateRows
            .map(row => sheet.getRange(row, estimateColumn).getA1Notation());
        if (invalidEstimateNotations.length) {
            sheet.getRangeList(invalidEstimateNotations).setBackground('#FFCCCB');
        }
        const skipWeekends = (date) => {
            while (date.getDay() === 0 || date.getDay() === 6) {
                date = new Date(date.getTime() + 24 * 3600 * 1000);
            }
            return date;
        };
        const scheduleStart = skipWeekends(ScheduleSettings.start);
        const startsRangeValues = Utils.arrayOf(startsRange.getHeight(), ['']);
        const endsRangeValues = Utils.arrayOf(endsRange.getHeight(), ['']);
        for (const [teamId, teamDayEstimates] of allTeamDaysEstimates.entries()) {
            const lanes = new Lanes(Team.getById(teamId).lanes);
            teamDayEstimates.forEach(dayEstimate => lanes.add(dayEstimate.daysEstimate, dayEstimate, laneIndex => dayEstimate.laneIndex = laneIndex));
            for (const lane of lanes.lanes) {
                let lastEnd = undefined;
                for (const dayEstimate of lane.objects()) {
                    let start = scheduleStart;
                    if (lastEnd != null) {
                        start = skipWeekends(new Date(lastEnd.getTime() + 24 * 3600 * 1000));
                    }
                    startsRangeValues[dayEstimate.index] = [start];
                    const daysEstimate = Math.ceil(dayEstimate.daysEstimate * (1 + ScheduleSettings.bufferCoefficient));
                    const end = lastEnd = skipWeekends(new Date(start.getTime() + daysEstimate * 24 * 3600 * 1000));
                    endsRangeValues[dayEstimate.index] = [end];
                }
            }
        }
        startsRange.setValues(startsRangeValues);
        endsRange.setValues(endsRangeValues);
        if (lanesRange != null) {
            const laneRangeValues = Utils.arrayOf(lanesRange.getHeight(), ['']);
            for (const y of Utils.range(1, lanesRange.getHeight())) {
                const index = y - 1;
                const dayEstimate = allDaysEstimates.find(it => it.index === index);
                if ((dayEstimate === null || dayEstimate === void 0 ? void 0 : dayEstimate.laneIndex) != null) {
                    laneRangeValues[index] = [`${dayEstimate.teamId}-${dayEstimate.laneIndex + 1}`];
                }
            }
            lanesRange.setValues(laneRangeValues);
        }
    }
}
class ScheduleSettings {
    static get start() {
        const settings = Settings.getMap(GSheetProjectSettings.settingsScheduleScope);
        const stringValue = settings.get('start');
        if (!(stringValue === null || stringValue === void 0 ? void 0 : stringValue.length)) {
            return new Date();
        }
        return new Date(stringValue);
    }
    static get bufferCoefficient() {
        const settings = Settings.getMap(GSheetProjectSettings.settingsScheduleScope);
        const stringValue = settings.get('bufferCoefficient');
        const value = parseFloat(stringValue !== null && stringValue !== void 0 ? stringValue : '');
        return isNaN(value) ? 0 : value;
    }
}
class Settings {
    static getMatrix(settingsScope) {
        const settingsSheet = SheetUtils.getSheetByName(GSheetProjectSettings.settingsSheetName);
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
            for (const row of Utils.range(scopeRow + 2, settingsSheet.getLastRow())) {
                const item = new Map();
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
            }
            return result;
        });
    }
    static getMap(settingsScope) {
        const settingsSheet = SheetUtils.getSheetByName(GSheetProjectSettings.settingsSheetName);
        settingsScope = Utils.normalizeName(settingsScope);
        return ExecutionCache.getOrComputeCache(['settings', 'map', settingsScope], () => {
            const scopeRow = this._findScopeRow(settingsSheet, settingsScope);
            if (scopeRow == null) {
                throw new Error(`Settings with "${settingsScope}" can't be found`);
            }
            const result = new Map();
            for (const row of Utils.range(scopeRow + 1, settingsSheet.getLastRow())) {
                const values = settingsSheet.getRange(row, 1, 1, 2).getValues()[0];
                const key = Utils.toLowerCamelCase(values[0].toString().trim());
                if (!key.length) {
                    break;
                }
                const value = values[1].toString().trim();
                result.set(key, value);
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
        if (sheet == null) {
            return undefined;
        }
        columnName = Utils.normalizeName(columnName);
        return ExecutionCache.getOrComputeCache(['findColumnByName', sheet, columnName], () => {
            for (const col of Utils.range(1, sheet.getLastColumn())) {
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
}
class State {
    static isStructureChanged() {
        const timestamp = this._loadStateTimestamp('lastStructureChange');
        return timestamp != null && this._now < timestamp;
    }
    static updateLastStructureChange() {
        this._now = new Date().getTime();
        this._saveStateTimestamp('lastStructureChange', this._now);
    }
    static reset() {
        this._now = new Date().getTime();
    }
    static _loadStateTimestamp(key) {
        var _a;
        const cache = CacheService.getDocumentCache();
        if (cache == null) {
            return null;
        }
        const timestamp = parseInt((_a = cache.get(`state:${key}`)) !== null && _a !== void 0 ? _a : '');
        return isNaN(timestamp) ? null : timestamp;
    }
    static _saveStateTimestamp(key, timestamp) {
        var _a;
        (_a = CacheService.getDocumentCache()) === null || _a === void 0 ? void 0 : _a.put(`state:${key}`, timestamp.toString());
    }
}
State._now = new Date().getTime();
class Team {
    static getAll() {
        var _a, _b, _c;
        const result = [];
        for (const info of Settings.getMatrix(GSheetProjectSettings.settingsTeamsScope)) {
            const id = (_b = (_a = info.get('id')) !== null && _a !== void 0 ? _a : info.get('teamId')) !== null && _b !== void 0 ? _b : info.get('team');
            if (!(id === null || id === void 0 ? void 0 : id.length)) {
                continue;
            }
            let lanes = parseInt((_c = info.get('lanes')) !== null && _c !== void 0 ? _c : '0');
            if (isNaN(lanes)) {
                lanes = 0;
            }
            result.push(new Team(id, lanes));
        }
        return result;
    }
    static findById(id) {
        return this.getAll().find(team => team.id === id);
    }
    static getById(id) {
        var _a;
        return (_a = this.findById(id)) !== null && _a !== void 0 ? _a : (() => {
            throw new Error(`"${id}" team can't be found`);
        })();
    }
    constructor(id, lanes) {
        this.id = id;
        this.lanes = Math.max(0, lanes);
    }
}
class Utils {
    static entryPoint(action) {
        try {
            State.reset();
            ExecutionCache.resetCache();
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
    static toLowerCamelCase(value) {
        value = value.replace(/^[^a-z0-9]+/i, '').replace(/[^a-z0-9]+$/i, '');
        if (value.length <= 1) {
            return value.toLowerCase();
        }
        value = value.substring(0, 1).toLowerCase() + value.substring(1).toLowerCase();
        value = value.replaceAll(/[^a-z0-9]+([a-z0-9])/ig, (_, letter) => letter.toUpperCase());
        return value;
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
