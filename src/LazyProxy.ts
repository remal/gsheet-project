class LazyProxy {

    static create<T>(supplier: () => T): T {
        const lazy = new Lazy<any>(supplier)
        const proxy = new Proxy({}, {
            apply(_, thisArg: any, argArray: any[]): any {
                const instance = lazy.get()
                return Reflect.apply(instance, thisArg, argArray)
            },
            construct(_, argArray: any[], newTarget: Function): object {
                const instance = lazy.get()
                return Reflect.construct(instance, argArray, newTarget)
            },
            defineProperty(_, property: string | symbol, attributes: PropertyDescriptor): boolean {
                const instance = lazy.get()
                return Reflect.defineProperty(instance, property, attributes)
            },
            deleteProperty(_, property: string | symbol): boolean {
                const instance = lazy.get()
                return Reflect.deleteProperty(instance, property)
            },
            get(_, property: string | symbol): any {
                const instance = lazy.get()
                return Reflect.get(instance, property, instance)
            },
            getOwnPropertyDescriptor(_, property: string | symbol): PropertyDescriptor | undefined {
                const instance = lazy.get()
                return Reflect.getOwnPropertyDescriptor(instance, property)
            },
            getPrototypeOf(_): object | null {
                const instance = lazy.get()
                return Reflect.getPrototypeOf(instance)
            },
            has(_, property: string | symbol): boolean {
                const instance = lazy.get()
                return Reflect.has(instance, property)
            },
            isExtensible(_): boolean {
                const instance = lazy.get()
                return Reflect.isExtensible(instance)
            },
            ownKeys(_): ArrayLike<string | symbol> {
                const instance = lazy.get()
                return Reflect.ownKeys(instance)
            },
            preventExtensions(_): boolean {
                const instance = lazy.get()
                return Reflect.preventExtensions(instance)
            },
            set(_, property: string | symbol, newValue: any): boolean {
                const instance = lazy.get()
                return Reflect.set(instance, property, newValue, instance)
            },
            setPrototypeOf(_, value: object | null): boolean {
                const instance = lazy.get()
                return Reflect.setPrototypeOf(instance, value)
            },
        })
        return proxy as T
    }

}
