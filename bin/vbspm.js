var {parseWSF, parseWSFStr} = require('wsf2json')
const fs = require('fs');
const { strict } = require('assert');


String.prototype.htmlEscape = function htmlEscape(str) {
    const CHAR_AMP = '&amp;'
    const CHAR_SINGLE = '&apos;';
    if (!str) { str = this; }
    return str.replace(CHAR_AMP, '&')
        .replace(CHAR_SINGLE, '\'');
};

parseWSF('bin\\build.wsf').then((jobs)=>{
    // console.log(JSON.stringify(jobs, null, 2));
    let vbsCombined = jobs.reduce((vbs, job)=>{
        let { id, script, runtime } = job;
        if (id) {
            vbs += `\r\n\r\n\r\n' ================================== Job: ${id} ================================== \r\n`
        }
        if (script) {
            vbs += script.reduce((s, scr)=>{
                let {type, src, exists, language, value} = scr;
                if (type) {
                    s += `\r\n' ================= ${type}`
                    if (type === 'src') {
                        s += ` : ${src}`
                    }
                    s += ` ================= \r\n`
                }
                if (language.toLowerCase() === "vbscript" && value) {
                    s += value;
                }
                return s;
            }, '');
        }
        //Inject arguments usage
        if (runtime) {
            let usage = runtime.reduce((str, param)=>{
                let {name, helpstring} = param;
                str += `Wscript.Echo "/${name}:  ${helpstring}"\r\n`;
                return str;
            },'');
            vbs = vbs.replace('WScript.Arguments.ShowUsage', usage);
        }
        return vbs;
    },'');
    vbsCombined = vbsCombined.htmlEscape();
    // console.log('vbs combined:')
    // console.log(vbsCombined);
    fs.writeFileSync('bin\\build.out.vbs', vbsCombined);
}).catch((error)=>{
    console.error(error)
})