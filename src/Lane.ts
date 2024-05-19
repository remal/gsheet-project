class Lane<T> {

    private readonly _amounts: number[] = []
    private readonly _objects: T[] = []

    add(amount: number, object: T) {
        this._amounts.push(amount)
        this._objects.push(object)
        return this
    }

    get sum(): number {
        return this._amounts.reduce((sum, amount) => sum + amount, 0)
    }

    * [Symbol.iterator](): Iterable<[number, T]> {
        for (let i = 0; i < this._amounts.length; ++i) {
            yield [this._amounts[i], this._objects[i]]
        }
    }

    * amounts(): Iterable<number> {
        for (const item of this._amounts) {
            yield item
        }
    }

    * objects(): Iterable<T> {
        for (const item of this._objects) {
            yield item
        }
    }

}
