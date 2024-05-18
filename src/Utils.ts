class Utils {

    static entryPoint<T>(action: () => T): T {
        try {
            return action()

        } catch (e) {
            console.error(e)
            throw e
        }
    }

    static* range(startIncluding: number, endIncluding: number): Generator<any, any, number> {
        for (let n = startIncluding; n <= endIncluding; ++n) {
            yield n
        }
    }

    static normalizeName(name: string): string {
        return name.toString().trim().replaceAll(/\s+/g, ' ').toLowerCase()
    }

    static extractRegex(string: string, regexp: string | RegExp, group?: number | string): string | null {
        if (this.isString(regexp)) {
            regexp = new RegExp(regexp)
        }

        if (group == null) {
            group = 0
        }

        const match = regexp.exec(string)
        if (match == null) {
            return null
        }

        if (match.groups != null) {
            const result = match.groups[group]
            if (result != null) {
                return result
            }
        }

        return match[group]
    }

    static distinct<T>(): (value: T, index: number, array: T[]) => boolean {
        return (value, index, array) => array.indexOf(value) === index
    }

    static distinctBy<T>(getter: (value: T) => any): (value: T, index: number, array: T[]) => boolean {
        const seen = new Set<any>()
        return (value) => {
            const property = getter(value)
            if (seen.has(property)) {
                return false
            }

            seen.add(property)
            return true
        }
    }

    static merge<T extends Record<string, any>, P extends Partial<T>>(...objects: P[]): T {
        const result = {}
        for (const object of objects) {
            if (object == null) {
                continue
            }

            for (const key of Object.keys(object)) {
                const value = object[key] as any
                if (value === undefined) {
                    continue
                }

                const currentValue = result[key]
                if (this.isRecord(value) && this.isRecord(currentValue)) {
                    result[key] = this.merge(currentValue, value)
                }

                result[key] = value
            }
        }
        return result as any as T
    }

    static arrayEquals<T>(array1?: T[], array2?: T[]): boolean {
        if (array1 === array2) {
            return true
        } else if (array1 == null) {
            return false
        } else if (array2 == null) {
            return false
        }

        if (array1.length !== array2.length) {
            return false
        }

        for (let i = 0; i < array1.length; ++i) {
            const element1 = array1[i]
            const element2 = array2[i]
            if (element1 !== element2) {
                return false
            }
        }

        return true
    }

    static isString(value: unknown): value is string {
        return typeof value === 'string'
    }

    static isFunction(value: unknown): value is Function {
        return typeof value === 'function'
    }

    static isRecord(value: unknown): value is Record<string, any> {
        return typeof value === 'object' && !Array.isArray(value)
    }

    static throwNotConfigured<T>(name: string): T {
        throw new Error(`Not configured: ${name}`)
    }

}
