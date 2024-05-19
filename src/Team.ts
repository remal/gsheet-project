class Team {

    static getAllTeams(): Team[] {
        const result: Team[] = []
        for (const info of Settings.getMatrix('teams')) {
            const id = info.get('id') ?? info.get('teamId') ?? info.get('team')
            if (!id.length) {
                continue
            }

            let lanes = parseInt(info.get('lanes'))
            if (isNaN(lanes)) {
                lanes = 0
            }

            result.push(new Team(id, lanes))
        }
        return result
    }


    id: string
    lanes: number

    constructor(id: string, lanes: number) {
        this.id = id
        this.lanes = Math.max(0, lanes)
    }

}
