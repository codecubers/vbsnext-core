'use strict';

//Async method (Windows):

const { spawn } = require( 'child_process' );
// NOTE: Windows Users, this command appears to be differ for a few users.
// You can think of this as using Node to execute things in your Command Prompt.
// If `cmd` works there, it should work here.
// If you have an issue, try `dir`:
// const dir = spawn( 'dir', [ '.' ] );
// const dir = spawn( 'cmd', [ '/c', 'dir' ] );

// dir.stdout.on( 'data', ( data ) => console.log( `stdout: ${ data }` ) );
// dir.stderr.on( 'data', ( data ) => console.log( `stderr: ${ data }` ) );
// dir.on( 'close', ( code ) => console.log( `child process exited with code ${code}` ) );
const winCommandAsync = function(cmd) {
    return new Promise((resolve, reject)=>{    
        let output = '';
        let error = '';
        let errCnt = 0;
        let errorType = '';
        try {
            const spawnned = spawn( 'cmd', [ '/c', cmd ] );
            spawnned.stdout.on( 'data', ( data ) => output += data );
            spawnned.stderr.on( 'data', ( data ) => {
                if (data.includes('Microsoft VBScript runtime error')) {
                    errorType = 'Runtime'
                }
                errCnt++
                error += data 
            });
            spawnned.on( 'close', ( code ) => {
                //console.log( `child process exited with code ${code}` ) 
                resolve({
                    code: code + errCnt, 
                    output, 
                    error,
                    errorType
                })
            });
        } catch (error) {
            reject(error)
        }
    })
}
module.exports = winCommandAsync



//Async method (Unix):

// 'use strict';

// const { spawn } = require( 'child_process' );
// const ls = spawn( 'ls', [ '-lh', '/usr' ] );

// ls.stdout.on( 'data', ( data ) => {
//     console.log( `stdout: ${ data }` );
// } );

// ls.stderr.on( 'data', ( data ) => {
//     console.log( `stderr: ${ data }` );
// } );

// ls.on( 'close', ( code ) => {
//     console.log( `child process exited with code ${ code }` );
// } );


//Sync:

// 'use strict';

// const { spawnSync } = require( 'child_process' );
// const ls = spawnSync( 'ls', [ '-lh', '/usr' ] );

// console.log( `stderr: ${ ls.stderr.toString() }` );
// console.log( `stdout: ${ ls.stdout.toString() }` );

