Dim pu
set pu = new PathUtil

Function test
    EchoX "0) \. => %x", pu.Resolve("\.")
    EchoX "1) . => %x", pu.Resolve(".")
    EchoX "2) .\ => %x", pu.Resolve(".\")
    EchoX "3) .. => %x", pu.Resolve("..")
    EchoX "4) ..\ => %x", pu.Resolve("..\")
    EchoX "5) ..\.\ => %x", pu.Resolve("..\.\")
    EchoX "6) ..\..\ => %x", pu.Resolve("..\..\")
    EchoX "7) ..\.\.\..\ => %x", pu.Resolve("..\.\.\..\")
    EchoX "8) PathUtil.vbs => %x", pu.Resolve("PathUtil.vbs")
    EchoX "9) .\PathUtil.vbs => %x", pu.Resolve(".\PathUtil.vbs")
    EchoX "10) .\.\PathUtil.vbs => %x", pu.Resolve(".\.\PathUtil.vbs")
    EchoX "11) pkg\pkg.vbs => %x", pu.Resolve("pkg\pkg.vbs")
    EchoX "12) .\pkg\pkg.vbs => %x", pu.Resolve(".\pkg\pkg.vbs")
    EchoX "13) C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\PathUtil\pkg\pkg.vbs => %x", pu.Resolve("C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\PathUtil\pkg\pkg.vbs")
    EchoX "14) pkg1\pkg1.vbs => %x", pu.Resolve("pkg1\pkg1.vbs")
    EchoX "15) .\pkg1\pkg1.vbs => %x", pu.Resolve(".\pkg1\pkg1.vbs")
    EchoX "16) C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\ArrayUtil\pkg1\pkg1.vbs => %x", pu.Resolve("C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\ArrayUtil\pkg1\pkg1.vbs")
End Function

Echo "Hellow"
EchoX "Hellow %x", "World"
EchoX "Hellow %x here comes %x", Array("World", "VbScript")
EchoD "Debug Only: Hellow"
EchoDX "Debug Only: Hellow %x", "World"
EchoDX "Debug Only: Hellow %x here comes %x", Array("World", "VbScript")

EchoX "Default ScriptPath => %x", pu.ScriptPath
EchoX "Default BasePath => %x", pu.BasePath
EchoX "Default Last TempBasePath => %x", pu.TempBasePath
' test

Echo "Setting BasePath to C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\PathUtil"
pu.BasePath = "C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\PathUtil"
EchoX "BasePath => %x", pu.BasePath
' test

Echo "Setting TempBasePath to C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\PathUtil\"
pu.TempBasePath = "C:\Users\nanda\git\xps.local.npm\vbspm\lib\core\PathUtil\"
EchoX "Last TempBasePath => %x", pu.TempBasePath

Echo "Setting TempBasePath to ..\ArrayUtil"
pu.TempBasePath = "..\ArrayUtil"
EchoX "Last TempBasePath => %x", pu.TempBasePath
' test

' Echo "Setting BasePath to ..\..\..\build"
' pu.BasePath = "..\..\..\build"
' EchoX "BasePath => %x", pu.BasePath
' test

' Echo "Setting BasePath to ..\..\build"
' pu.BasePath = "..\..\build"
' EchoX "BasePath => %x", pu.BasePath
' test

' Echo "Setting BasePath to .\"
' pu.BasePath = ".\"
' EchoX "BasePath => %x", pu.BasePath
' test