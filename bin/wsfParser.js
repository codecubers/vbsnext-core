var {parseWSF, parseWSFStr, extractVBS} = require('wsf2json')
const fs = require('fs');
const path = require('path');
const { strict } = require('assert');


String.prototype.htmlEscape = function htmlEscape(str) {
    const CHAR_AMP = '&amp;'
    const CHAR_SINGLE = '&apos;';
    if (!str) { str = this; }
    return str.replace(CHAR_AMP, '&')
        .replace(CHAR_SINGLE, '\'');
};
let fBuilder = path.join("bin", "builder.wsf");
let fBuilderOut = path.join("bin", "vbsnext-build-unresolved.vbs");
parseWSF(fBuilder).then((jobs)=>{
    // console.log(JSON.stringify(jobs, null, 2));
    let vbsCombined = extractVBS(jobs);
    vbsCombined = vbsCombined.htmlEscape();
    // console.log('vbs combined:')
    // console.log(vbsCombined);
    fs.writeFileSync(fBuilderOut, vbsCombined);
}).catch((error)=>{
    console.error(error)
})
let fRunner = path.join("bin", "runner.wsf");
let fRunnerOut = path.join("bin", "vbsnext-unresolved.vbs");
parseWSF(fRunner).then((jobs)=>{
    // console.log(JSON.stringify(jobs, null, 2));
    let vbsCombined = extractVBS(jobs);
    vbsCombined = vbsCombined.htmlEscape();
    // console.log('vbs combined:')
    // console.log(vbsCombined);
    fs.writeFileSync(fRunnerOut, vbsCombined);
}).catch((error)=>{
    console.error(error)
})