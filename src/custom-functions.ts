/**
 * SHA-256 digest of the provided input
 * @param {unknown} value
 * @returns {string}
 * @customFunction
 */
function SHA256(value: unknown): string {
    const string = value?.toString() ?? ''
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, string)
    return Utilities.base64EncodeWebSafe(digest)
}
