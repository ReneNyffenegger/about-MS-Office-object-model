option explicit

sub main() ' {

    dim shFoo, shBar, shBaz as workSheet

    set shFoo = createSheet("foo")
    set shBar = createSheet("baz")
    set shBaz = createSheet("bar")

  '
  ' Using copy without arguments creates a
  ' new workbook that contains the copied Worksheet object.
  ' The copied worksheet retains
  '   - its name
  '   - its codeName property
  '
    shBar.copy

  dim newWorkbook as workbook
  set newWorkbook = application.activeWorkbook
  newWorkbook.sheets("Baz").cells(3,2) = "Copied sheet"

end sub ' }

function createSheet(name as string) as workSheet ' {
    set createSheet = activeWorkbook.sheets.add
    createSheet.name = name
    createSheet.cells(2,2) = "This is sheet " & name
end function ' }
