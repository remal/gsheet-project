class Formulas {

    static processFormula(formula: string): string {
        formula = formula.replaceAll(/#SELF_COLUMN\(([^)]+)\)/g, 'INDIRECT("RC"&COLUMN($1), FALSE)')
        formula = formula.replaceAll(/#SELF(\b|&)/g, 'INDIRECT("RC", FALSE)$1')
        return formula.split(/[\r\n]+/)
            .map(line => line.replace(/^\s+/, ''))
            .filter(line => line.length)
            .map(line => line.replaceAll(/^([<>&=*/+-]+ )/g, ' $1'))
            .map(line => line.replaceAll(/\s*\t\s*/g, ' '))
            .map(line => line.replaceAll(/"\s*&\s*""/g, '"'))
            .map(line => line.replaceAll(/([")])\s*&\s*([")])/g, '$1 & $2'))
            .map(line => line + (line.endsWith(',') || line.endsWith(';') ? ' ' : ''))
            .join('')
            .trim()
    }

    static addFormulaMarker(formula: string, marker: string | null | undefined): string {
        if (!marker?.length) {
            return formula
        }

        formula = formula.replace(/^=+/, '')
        formula = `IF("GSPf"<>"${marker}", ${formula}, "")`
        return '=' + formula
    }

    static addFormulaMarkers(formula: string, ...markers: (string | null | undefined)[]): string {
        markers = markers.filter(it => it?.length)
        if (!markers?.length) {
            return formula
        }

        if (markers.length === 1) {
            return this.addFormulaMarker(formula, markers[0])
        }

        formula = formula.replace(/^=+/, '')
        formula = `IF(AND(${markers.map(marker => `"GSPf"<>"${marker}"`).join(', ')}), ${formula}, "")`
        return '=' + formula
    }

    static extractFormulaMarkers(formula: string | null | undefined): string[] {
        if (!formula?.length) {
            return []
        }

        const markers = Utils.arrayOf<string>()
        const regex = /"GSPf"\s*<>\s*"([^"]+)"/g
        let match: RegExpExecArray | null
        while ((match = regex.exec(formula)) !== null) {
            markers.push(match[1])
        }
        return markers
    }

    static escapeFormulaString(string: string): string {
        return string.replaceAll(/"/g, '""')
    }

}
