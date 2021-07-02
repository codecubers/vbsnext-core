const fs = require('fs');
const extendVbs = require('vbs-method-parser')
let source = fs.readFileSync('.\\vbspm-bulk.vbs').toString();
extendVbs(source).then((resolved)=>{
    fs.writeFileSync('.\\vbspm.vbs', resolved);
    source = fs.readFileSync('.\\vbspm-build-bulk.vbs').toString();
    extendVbs(source).then((resolved)=>{
        fs.writeFileSync('.\\vbspm-build.vbs', resolved);
    })
})