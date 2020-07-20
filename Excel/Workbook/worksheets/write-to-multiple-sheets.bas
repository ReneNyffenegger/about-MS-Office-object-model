'
'      Compare with C:/Users/r.nyffenegger/github/github/notes/notes/Microsoft/Office/Excel/Object-Model/Workbook/exportAsFixedFormat
'
option explicit

sub main() ' {

    createSheet "sheet one"
    createSheet "sheet two"
    createSheet "sheet three"
    createSheet "sheet four"
    createSheet "sheet five"

    activeWorkbook.worksheets(array("sheet one", "sheet three", "sheet five")).select
    cells(2,1).select
    activeCell.value = "Odd"

    activeWorkbook.worksheets(array("sheet two", "sheet four")).select
    cells(2,1).select
    activeCell.value = "Even"

end sub ' }

sub createSheet(name as string) ' {

    dim sh as worksheet
    set sh = activeWorkbook.worksheets.add
    sh.name = name

    sh.cells(1,1) = "This is sheet " & name

end sub ' }
