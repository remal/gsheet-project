/**
 * SHA-256 digest of the provided input
 * @param {unknown} value
 * @returns {string}
 * @customFunction
 */
function SHA256(value: unknown): string {
    const string = value?.toString() ?? ''
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, string)
    return digest
        .map(num => num < 0 ? num + 256 : num)
        .map(num => num.toString(16))
        .map(num => (num.length === 1 ? '0' : '') + num)
        .join('')
}
