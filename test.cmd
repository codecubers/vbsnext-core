call cscript //nologo bin\vbsnext.vbs /file:\"test\test.vbs\"
call node bin\vbsnext-cls-resolver.js /file:test\test.vbs
call cscript //nologo build\test-bundle.vbs