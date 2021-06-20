Dim xl
set xl = new Excel
xl.OpenWorkBook(".\Excel_MVC_Creator.xlsm")
EchoX "Active workbook name is: %x", xl.GetActiveWorkbook.Name
xl.ExportVBComponents(".\Code")
xl.CloseWorkBook
set xl = nothing