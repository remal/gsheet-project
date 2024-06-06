class Timer {

    private readonly _name: string
    private readonly _start: number

    constructor(name: string) {
        this._name = name
        this._start = Date.now()
    }

    log() {
        console.log(`${this._name}: ${Date.now() - this._start}ms`)
    }

}
