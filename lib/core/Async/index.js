// var AsyncVbs = require('./Async-Vbs')
// AsyncVbs(function(out){
//     console.log(out)
// })

// const system = require('system-commands')
// system('dir').then(output => {
//     console.log(output)
// }).catch(error => {
//     console.error(error)
// })

const winCommandAsync = require('./Spawn-Vbs')
winCommandAsync('cscript //nologo build\\test-bundle.vbs').then((data)=>{
    let { code, output, error, errorType } = data;
    console.log("code: " + code)
    if (code == 0) {
        console.log("Command execution Successfully executed.")
        console.log(output)
    } else {
        console.log((errorType ? errorType : "Unknown") + " error occurred in exeuction.")
        console.log(output + "\r\n" + error)
    }
}).catch(error => console.log(error));