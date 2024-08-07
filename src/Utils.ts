class Utils {

    static* range(startIncluding: number, endIncluding: number): Iterable<number> {
        for (let n = startIncluding; n <= endIncluding; ++n) {
            yield n
        }
    }

    static normalizeName(name: string): string {
        return name.toString().trim().replaceAll(/\s+/g, ' ').toLowerCase()
    }

    static toLowerCamelCase(value: string): string {
        value = value.replace(/^[^a-z0-9]+/i, '').replace(/[^a-z0-9]+$/i, '')
        if (value.length <= 1) {
            return value.toLowerCase()
        }

        value = value.substring(0, 1).toLowerCase() + value.substring(1).toLowerCase()
        value = value.replaceAll(/[^a-z0-9]+([a-z0-9])/ig, (_, letter) => letter.toUpperCase())
        return value
    }

    static toUpperCamelCase(value: string): string {
        value = this.toLowerCamelCase(value)
        if (value.length <= 1) {
            return value.toUpperCase()
        }

        value = value.substring(0, 1).toUpperCase() + value.substring(1)
        return value
    }

    static mapRecordValues<V, VR>(
        record: Record<string, V>,
        transformer: (value: V, key: string) => VR,
    ): Record<string, VR> {
        const result = {} as Record<string, VR>
        Object.entries(record).forEach(([key, value]) => {
            result[key] = transformer(value, key)
        })
        return result
    }

    static mapToRecord<V>(keys: string[], transformer: (value: string) => V): Record<string, V> {
        const result = {} as Record<string, V>
        keys.forEach(key => {
            result[key] = transformer(key)
        })
        return result
    }

    /**
     * See https://stackoverflow.com/a/44134328/3740528
     */
    static hslToRgb(h: number, s: number, l: number): string {
        l /= 100
        const a = s * Math.min(l, 1 - l) / 100
        const f = (n: number) => {
            const k = (n + h / 30) % 12
            const color = l - a * Math.max(Math.min(k - 3, 9 - k, 1), -1)
            return Math.round(255 * color).toString(16).padStart(2, '0')
        }
        return `#${f(0)}${f(8)}${f(4)}`
    }

    static toJsonObject(object: any, callGetters: boolean = true, keepNulls: boolean = false): any {
        if (object == null) {
            return object

        } else if (object instanceof Date) {
            return new Date(object.getTime())

        } else if (Array.isArray(object)) {
            const result: any[] = []
            for (const element of object) {
                if (element === object || element === undefined || (!keepNulls && element === null)) {
                    continue
                }

                result.push(this.toJsonObject(element, callGetters))
            }
            return result

        } else if (this.isFunction(object.toJSON)) {
            return object.toJSON()

        } else if (this.isFunction(object.getA1Notation)) {
            return object.getA1Notation()

        } else if (this.isFunction(object.getSheetName)) {
            return object.getSheetName()

        } else if (typeof object === 'object') {
            const prototypePropertiesToExclude = ['constructor']

            const properties: string[] = []
            for (const property in object) {
                if (!object.hasOwnProperty(property)
                    && prototypePropertiesToExclude.includes(property)
                ) {
                    continue
                }

                properties.push(property)
            }
            properties.sort((p1, p2) => {
                const n1 = parseFloat(p1)
                const n2 = parseFloat(p2)
                if (!isNaN(n1) && !isNaN(n2)) {
                    return n1 - n2
                }

                return p1.localeCompare(p2)
            })

            const result: any = {}
            for (const property of properties) {
                const value = object[property]
                if (value === object || value === undefined || (!keepNulls && value === null)) {
                    continue
                }

                if (this.isFunction(value)) {
                    if (callGetters) {
                        const getterMatcher = property.match(/^(get|is)([A-Z].*)$/)
                        if (getterMatcher != null) {
                            const propValue = value.call(object)
                            if (propValue === object || propValue === undefined || (!keepNulls && propValue === null)) {
                                continue
                            }

                            let name = getterMatcher[2]
                            name = name.substring(0, 1).toLowerCase() + name.substring(1)
                            result[name] = this.toJsonObject(propValue, callGetters)
                        }
                    }
                    continue
                }

                result[property] = this.toJsonObject(value, callGetters)
            }
            return result

        } else {
            return object
        }
    }

    static groupBy<T>(array: T[], keyGetter: (element: T) => string | null | undefined): Map<string, T[]> {
        const result = new Map<string, T[]>()
        for (const element of array) {
            const key = keyGetter(element)
            if (key != null) {
                let groupedElements = result.get(key)
                if (groupedElements == null) {
                    groupedElements = []
                    result.set(key, groupedElements)
                }
                groupedElements.push(element)
            }
        }
        return result
    }

    static extractRegex(
        string: string | null | undefined,
        regexp: string | RegExp,
        group?: number | string,
    ): string | null {
        if (string == null) {
            return null;
        }

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

    static trimArrayEndBy<T>(array: T[], predicate: (element: T) => boolean) {
        while (array.length) {
            const lastElement = array[array.length - 1]
            if (predicate(lastElement)) {
                --array.length
            } else {
                break
            }
        }
    }

    static arrayRemoveIf<T>(array: T[], predicate: (element: T) => boolean) {
        for (let index = 0; index < array.length; ++index) {
            const element = array[index]
            if (predicate(element)) {
                array.splice(index, 1)
                --index
            }
        }
    }

    static moveArrayElements(array: any[], fromIndex: number, count: number, targetIndex: number) {
        if (fromIndex === targetIndex || count <= 0) {
            return
        }

        if (array.length <= targetIndex) {
            array.length = targetIndex + 1
        }

        const moved = array.splice(fromIndex, count)
        array.splice(targetIndex, 0, ...moved)
    }

    static parseDate(value: any): Date | null {
        if (value == null) {
            return null
        } else if (this.isNumber(value)) {
            return new Date(value)
        } else if (Utils.isString(value)) {
            try {
                return new Date(Number.isNaN(value) ? value : parseFloat(value))
            } catch (_) {
                return null
            }
        } else if (this.isFunction(value.getTime)) {
            return this.parseDate(value.getTime())
        } else {
            return null
        }
    }

    static parseDateOrThrow(value: any): Date {
        return this.parseDate(value) ?? (() => {
            throw new Error(`Not a date: "${value}"`)
        })()
    }

    static hashCode(value: string | null | undefined): number {
        if (!value?.length) {
            return 0
        }

        let hash: number = 0
        for (let i = 0; i < value.length; ++i) {
            const chr = value.charCodeAt(i)
            hash = ((hash << 5) - hash) + chr
            hash |= 0
        }
        return hash
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
                    continue
                }

                result[key] = value
            }
        }
        return result as any as T
    }

    static mergeInto<T extends Record<string, any>, P extends Partial<T>>(result: T, ...objects: P[]): T {
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
                    this.mergeInto(currentValue, value)
                    continue
                }

                (result as any)[key] = value
            }
        }

        return result
    }

    static arrayEquals<T>(array1: T[] | null | undefined, array2: T[] | null | undefined): boolean {
        if (array1 === array2) {
            return true
        } else if (array1 == null || array2 == null) {
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

    static arrayOf<T>(length?: number, initValue?: T): T[] {
        const array = new Array<T>(length ?? 0)
        if (initValue !== undefined) {
            array.fill(initValue)
        }
        return array
    }

    static escapeRegex(string: string): string {
        return string.replaceAll(/[.*+?^${}()|[\]\\]/g, '\\$&')
    }

    static numericAsc(): (n1: number, n2: number) => number {
        return (n1, n2) => n1 - n2
    }

    static numericDesc(): (n1: number, n2: number) => number {
        return (n1, n2) => n2 - n1
    }

    static isString(value: unknown): value is string {
        return typeof value === 'string'
    }

    static isNumber(value: unknown): value is number {
        return typeof value === 'number'
    }

    static isBoolean(value: unknown): value is boolean {
        return typeof value === 'boolean'
    }

    static isFunction(value: unknown): value is Function {
        return typeof value === 'function'
    }

    static isRecord(value: unknown): value is Record<string, any> {
        return typeof value === 'object' && !Array.isArray(value)
    }

    static throwNotImplemented<T>(...name: string[]): T {
        throw new Error(`Not implemented: ${name.join(': ')}`)
    }

}
