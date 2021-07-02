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
winCommandAsync('cscript //nologo build\\plotXY-bundle.vbs /file:src\plotXY.vbs /workbook:workbooks\SimpleXYPlot.xlsm /destination:ChartPlots /workbook:dummy /debug:true /data:A,B,1,1,2,4,3,9,4,0').then((data)=>{
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