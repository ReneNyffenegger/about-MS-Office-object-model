option explicit

sub main() ' {

    activeWorkbook.saveAs _
       fileName   := environ("TEMP") & "\" & "bla.xlsm" , _
       fileFormat := xlOpenXMLWorkbookMacroEnabled

end sub ' }

'
'  vim: ft=vb
'
