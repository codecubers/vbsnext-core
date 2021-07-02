let file = process.argv[2];
console.log(`file receive: ${file}`)
var filename = file.replace(/^.*[\\\/]/, '')
console.log('filename', filename)
file = filename.replace('.vbs', '-bundle-unresolved.vbs')
file = '.\\build\\' + file;
console.log('filename finally...', file)

const fs = require('fs');
const extendVbs = require('vbs-method-parser')
let source = fs.readFileSync(file).toString();
// console.log('source', source)
extendVbs(source).then((resolved)=>{
    let outFile = file.replace('-unresolved.vbs', '.vbs');
    console.log(`Writing resolved file to: ${outFile}`)
    fs.writeFileSync(outFile, resolved);
})