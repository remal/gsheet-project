const fs = require('fs')
const path = require('path')
const rimraf = require('rimraf')
const crypto = require('crypto')

const buildDir = 'build'
const mainClass = 'GSheetProject'
const outputFile = `dist/${mainClass}.js`

const distDir = path.dirname(outputFile)
rimraf.sync(distDir)
fs.mkdirSync(distDir, {recursive: true})

const files = fs.readdirSync(buildDir)
    .toSorted((f1, f2) => {
        if (f1.startsWith('_') && !f2.startsWith('_')) {
            return -1
        } else if (!f1.startsWith('_') && f2.startsWith('_')) {
            return 1
        }
        return f1.localeCompare(f2)
    })

const contents = []
const digest = crypto.createHash('sha256')
for (const file of files) {
    let content = fs.readFileSync(`${buildDir}/${file}`, {encoding: 'UTF-8'}).trim()
    content = content.replaceAll(/\/\*[^*][\s\S]*?\*\//g, '')
    contents.push(content)
    digest.update(content, 'utf-8')
}

let content = contents.join('\n')
const hash = digest.digest('hex')
content = content.replaceAll(/(['"`])([^'"`]*)\$\$\$HASH\$\$\$([^'"`]*)\1/g, `$1$2${hash}$3$1`)

fs.appendFileSync(outputFile, content, {encoding: 'UTF-8'})
