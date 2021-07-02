const fs = require('fs');
const extendVbs = require('vbs-method-parser')
let source = fs.readFileSync('.\\vbspm-unresolved.vbs').toString();
extendVbs(source).then((resolved)=>{
    fs.writeFileSync('.\\vbspm.vbs', resolved);
    source = fs.readFileSync('.\\vbspm-build-unresolved.vbs').toString();
    fs.unlinkSync('.\\vbspm-unresolved.vbs')
    extendVbs(source).then((resolved)=>{
        fs.writeFileSync('.\\vbspm-build.vbs', resolved);
        fs.unlinkSync('.\\vbspm-build-unresolved.vbs')
    })
})