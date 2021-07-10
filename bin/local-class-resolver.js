const fs = require('fs');
const path = require('path');
const extendVbs = require('@vbsnext/vbs-class-extends')

let fRunnerUnresolved = path.join("bin", "vbsnext-unresolved.vbs");
let fRunnerResolved = path.join("bin", "vbsnext.vbs");
let fBuilderUnresolved = path.join("bin", "vbsnext-build-unresolved.vbs");
let fBUilderResolved = path.join("bin", "vbsnext-build.vbs");

let source = fs.readFileSync(fRunnerUnresolved).toString();
extendVbs(source).then((resolved)=>{
    fs.writeFileSync(fRunnerResolved, resolved);
    source = fs.readFileSync(fBuilderUnresolved).toString();
    fs.unlinkSync(fRunnerUnresolved)
    extendVbs(source).then((resolved)=>{
        fs.writeFileSync(fBUilderResolved, resolved);
        fs.unlinkSync(fBuilderUnresolved)
    })
})