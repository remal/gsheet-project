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

    static isString(value: unknown): value is string {
        return typeof value === 'string'
    }

    static isFunction(value: unknown): value is Function {
        return typeof value === 'function'
    }

}
