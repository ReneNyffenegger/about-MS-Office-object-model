option explicit

sub main() ' {

    dim wb as workbook
    set wb = workbooks.add(xlWBATWorksheet)

    dim shTwo   as workSheet
    dim shOne   as workSheet
    dim shThree as workSheet

    set shTwo    =  wb.workSheets(1)
    shTwo.name   = "Two"

    set shOne    =  wb.workSheets.add ' Add to the left, by default
    shOne.name   = "One"

    set shThree  =  wb.workSheets.add(after := shTwo)
    shThree.name = "Three"

end sub ' }
