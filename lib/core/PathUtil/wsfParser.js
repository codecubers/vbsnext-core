var {parseWSF, parseWSFStr, extractVBS} = require('wsf2json')
const fs = require('fs');

String.prototype.htmlEscape = function htmlEscape(str) {
    const CHAR_AMP = '&amp;'
    const CHAR_SINGLE = '&apos;';
    if (!str) { str = this; }
    return str.replace(CHAR_AMP, '&')
        .replace(CHAR_SINGLE, '\'');
};

parseWSF('test.wsf').then((jobs)=>{
    // console.log(JSON.stringify(jobs, null, 2));
    let vbsCombined = extractVBS(jobs);
    vbsCombined = vbsCombined.htmlEscape();
    // console.log('vbs combined:')
    // console.log(vbsCombined);
    fs.writeFileSync('PathUtil.test.vbs', vbsCombined);
}).catch((error)=>{
    console.error(error)
})
