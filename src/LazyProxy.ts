class LazyProxy {

    private static readonly _lazyProxyToLazy = new WeakMap<any, Lazy<any>>()

    static create<T>(supplier: () => T): T {
        const lazy = new Lazy<any>(supplier)
        const proxy = new Proxy(lazy, {
            apply(lazy, thisArg: any, argArray: any[]): any {
                const instance = lazy.get()
                argArray = argArray.map(it => this.unwrapLazyProxy(it))
                return Reflect.apply(instance, thisArg, argArray)
            },
            construct(lazy, argArray: any[], newTarget: Function): object {
                const instance = lazy.get()
                argArray = argArray.map(it => this.unwrapLazyProxy(it))
                return Reflect.construct(instance, argArray, newTarget)
            },
            defineProperty(lazy, property: string | symbol, attributes: PropertyDescriptor): boolean {
                const instance = lazy.get()
                return Reflect.defineProperty(instance, property, attributes)
            },
            deleteProperty(lazy, property: string | symbol): boolean {
                const instance = lazy.get()
                return Reflect.deleteProperty(instance, property)
            },
            get(lazy, property: string | symbol, receiver: any): any {
                const instance = lazy.get()
                let value = Reflect.get(instance, property, instance)
                if (Utils.isFunction(value)) {
                    return function () {
                        const target = this === receiver ? instance : this
                        const argArray = Array.from(arguments).map(it => LazyProxy.unwrapLazyProxy(it))
                        return value.apply(target, argArray)
                    }
                }
                return value
            },
            getOwnPropertyDescriptor(lazy, property: string | symbol): PropertyDescriptor | undefined {
                const instance = lazy.get()
                return Reflect.getOwnPropertyDescriptor(instance, property)
            },
            getPrototypeOf(lazy): object | null {
                const instance = lazy.get()
                return Reflect.getPrototypeOf(instance)
            },
            has(lazy, property: string | symbol): boolean {
                const instance = lazy.get()
                return Reflect.has(instance, property)
            },
            isExtensible(lazy): boolean {
                const instance = lazy.get()
                return Reflect.isExtensible(instance)
            },
            ownKeys(lazy): ArrayLike<string | symbol> {
                const instance = lazy.get()
                return Reflect.ownKeys(instance)
            },
            preventExtensions(lazy): boolean {
                const instance = lazy.get()
                return Reflect.preventExtensions(instance)
            },
            set(lazy, property: string | symbol, newValue: any): boolean {
                const instance = lazy.get()
                return Reflect.set(instance, property, newValue, instance)
            },
            setPrototypeOf(lazy, value: object | null): boolean {
                const instance = lazy.get()
                return Reflect.setPrototypeOf(instance, value)
            },
        })
        this._lazyProxyToLazy.set(proxy, lazy)
        return proxy as T
    }

    static unwrapLazyProxy<T>(lazyProxy: T): T {
        const lazy = this._lazyProxyToLazy.get(lazyProxy)
        if (lazy == null) {
            return lazyProxy
        }

        return lazy.get() as T
    }

}
