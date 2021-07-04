const fs = require('fs');
const extendVbs = require('@vbsnext/vbs-class-extends')
let source = fs.readFileSync('bin\\vbsnext-unresolved.vbs').toString();
extendVbs(source).then((resolved)=>{
    fs.writeFileSync('bin\\vbsnext.vbs', resolved);
    source = fs.readFileSync('bin\\vbsnext-build-unresolved.vbs').toString();
    fs.unlinkSync('bin\\vbsnext-unresolved.vbs')
    extendVbs(source).then((resolved)=>{
        fs.writeFileSync('bin\\vbsnext-build.vbs', resolved);
        fs.unlinkSync('bin\\vbsnext-build-unresolved.vbs')
    })
})