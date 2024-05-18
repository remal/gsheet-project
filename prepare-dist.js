const fs = require('fs')
const path = require('path')
const rimraf = require('rimraf')

const buildDir = 'build'
const mainClass = 'GSheetProject'
const outputFile = `dist/${mainClass}.js`

const distDir = path.dirname(outputFile)
rimraf.sync(distDir)
fs.mkdirSync(distDir, {recursive: true})

const files = fs.readdirSync(buildDir)
    .sort((f1, f2) => {
        if (f1.startsWith(mainClass) && !f2.startsWith(mainClass)) {
            return -1
        } else if (!f1.startsWith(mainClass) && f2.startsWith(mainClass)) {
            return 1
        }
        return f1.localeCompare(f2)
    })

for (const file of files) {
    const content = fs.readFileSync(`${buildDir}/${file}`)
    fs.appendFileSync(outputFile, content)
}
