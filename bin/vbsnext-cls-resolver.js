let file = process.argv[2];
console.log(`file receive: ${file}`)
var filename = file.replace(/^.*[\\\/]/, '')
console.log('filename', filename)
file = filename.replace('.vbs', '-bundle-unresolved.vbs')
file = '.\\build\\' + file;
console.log('filename finally...', file)
let outFile = file.replace('-unresolved.vbs', '.vbs');

const fs = require('fs');
const extendVbs = require('@vbsnext/vbs-class-extends')
let source = fs.readFileSync(file).toString();
extendVbs(source).then((resolved)=>{
    console.log(`Writing resolved file to: ${outFile}`)
    fs.writeFileSync(outFile, resolved)
    console.log('Deleting unresolved file')
    fs.unlinkSync(file);
})