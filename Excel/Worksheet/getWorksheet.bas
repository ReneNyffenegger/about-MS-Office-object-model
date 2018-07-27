function getWorksheet(name_ as string) as excel.worksheet
 '
 '  Return worksheet with the given name.
 '  If it doesn't exist, it is created.
 '

    on error goto createWorksheet
       set getWorksheet = thisWorkbook.sheets(name_)
    '  Worksheet exists, we can return the function:
       exit function

    createWorksheet:
    '  Error encountered, we have to create the worksheet

       set getWorksheet = thisWorkbook.sheets.add(after := thisWorkbook.sheets(thisWorkbook.sheets.count))
           getWorksheet.name = name_

end function

sub main() ' {
    dim ws as worksheet

    set ws = getWorksheet("tq84")
    ws.cells(1,1) = "Hello world"

end sub ' }
