class Lazy<T> {

    private value: T
    private supplier: () => T

    constructor(supplier: () => T) {
        this.supplier = supplier
    }

    get(): T {
        if (this.supplier != null) {
            this.value = this.supplier()
            this.supplier = null
        }

        return this.value
    }

}
