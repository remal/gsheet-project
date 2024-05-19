class Team {

    static getAll(): Team[] {
        const result: Team[] = []
        for (const info of Settings.getMatrix(GSheetProjectSettings.settingsTeamsScope)) {
            const id = info.get('id') ?? info.get('teamId') ?? info.get('team')
            if (!id?.length) {
                continue
            }

            let lanes = parseInt(info.get('lanes') ?? '0')
            if (isNaN(lanes)) {
                lanes = 0
            }

            result.push(new Team(id, lanes))
        }
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

    constructor(id: string, lanes: number) {
        this.id = id
        this.lanes = Math.max(0, lanes)
    }

}
