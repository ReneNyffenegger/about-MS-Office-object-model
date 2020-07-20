option explicit

sub main() ' {

    createSheet "sheet one"
    createSheet "sheet two"
    createSheet "sheet three"
    createSheet "sheet four"
    createSheet "sheet five"

    activeWorkbook.sheets(array("sheet two", "sheet five", "sheet four")).select
    
    activeSheet.exportAsFixedFormat           _
       type             :=  xlTypePDF       , _
       fileName         := "exported-sheets", _
       openafterpublish :=  true            , _
       ignoreprintareas :=  false

end sub ' }

sub createSheet(name as string) ' {

    dim sh as worksheet
    set sh = activeWorkbook.sheets.add
    sh.name = name

    sh.cells(2,2) = name

end sub ' }
