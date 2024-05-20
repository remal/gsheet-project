class Team {

    static getAll(): Team[] {
        const result: Team[] = []
        Settings.getMatrix(GSheetProjectSettings.settingsTeamsScope).forEach((info, index, allInfos) => {
            const id = info.get('id')
                ?? info.get('teamId')
                ?? info.get('team')
            if (!id?.length) {
                return
            }

            let lanes = parseInt(info.get('lanes') ?? '0')
            if (isNaN(lanes)) {
                lanes = 0
            }

            const color = info.get('color')
                ?? info.get('colour')
                ?? Utils.hslToRgb(360 * index / allInfos.length, 50, 80)

            const team = new Team(id, lanes, color)
            result.push(team)

            if ((info as any).hasOwnProperty('$settingsRange')) {
                if (State.isStructureChanged()) return
                const settingsRange = info['$settingsRange'] as SettingsRange
                Settings.settingsSheet.getRange(
                    settingsRange.row,
                    settingsRange.column,
                    settingsRange.rows,
                    settingsRange.columns,
                ).setBackground(team.color)
            }
        })
        return result
    }

    static findById(id: string): Team | undefined {
        return this.getAll().find(team => team.id === id)
    }

    static getById(id: string): Team {
        return this.findById(id) ?? (() => {
            throw new Error(`"${id}" team can't be found`)
        })()
    }


    readonly id: string
    readonly lanes: number
    readonly color: string

    constructor(id: string, lanes: number, color: string) {
        this.id = id
        this.lanes = Math.max(0, lanes)
        this.color = color
    }

}
