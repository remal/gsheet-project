class Lazy<T> {

    private _value: T
    private _supplier?: () => T

    constructor(supplier: () => T) {
        this._supplier = supplier
    }

    get(): T {
        if (this._supplier != null) {
            this._value = this._supplier()
            this._supplier = undefined
        }

        return this._value
    }

}
