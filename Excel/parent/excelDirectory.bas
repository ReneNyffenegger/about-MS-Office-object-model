option explicit

sub main() ' {

    activeWorkbook.saveAs _
       fileName   := environ("TEMP") & application.pathSeparator & "test.xlsm" , _
       fileFormat := xlOpenXMLWorkbookMacroEnabled

    cells(1,1).value   = "This workbook is in the directory:"
    cells(2,1).formula = "= worksheetDirectory(a1)"

end sub ' }

function worksheetDirectory(rng as range) as string ' {
    worksheetDirectory = rng.parent.parent.path
end function ' }
