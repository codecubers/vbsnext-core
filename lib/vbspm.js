var {parseWSF, parseWSFStr} = require('wsf2json')
const fs = require('fs');
const { strict } = require('assert');

parseWSF('vbspm.wsf').then((jobs)=>{
    console.log(JSON.stringify(jobs, null, 2));
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
    console.log('vbs combined:')
    console.log(vbsCombined);
    fs.writeFileSync('vbspm.out.vbs', vbsCombined);
}).catch((error)=>{
    console.error(error)
})