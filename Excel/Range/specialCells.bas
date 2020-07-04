option explicit

sub main() ' {

  '
  ' Create the date for the demonstration
  '
    testData

  '
  ' range.specialCells returns a sub-range. We use xlCellTypeFormulas to
  ' get the range that contains cells and assign the returned range
  ' to rngFormulas …
  '
    dim rngFormulas as range
    set rngFormulas = activesheet.cells.specialCells(xlCellTypeFormulas)
  '
  ' … and color the range in a light color:
  '
    rngFormulas.interior.color = rgb(255, 240, 230)

  '
  ' Of course, we can color a range without first assigning the returned
  ' range to a variable.
  '
  ' The following two lines color cells that contain a constant number (second argument = 1)
  ' with a violett color and cells that contain a constant text (second argument = 2)
  ' with a greenish color:
  '
    activesheet.cells.specialCells(xlCellTypeConstants, 1).interior.color = rgb(240, 230, 255)
    activesheet.cells.specialCells(xlCellTypeConstants, 2).interior.color = rgb(230, 255, 240)

    columns(1).autofit
    columns(3).autofit

end sub ' }

sub testData() ' {

    activeSheet.usedRange.clearContents
    activeSheet.usedRange.clearFormats

    cells(1,1).value     = "Fibonacci"
    cells(1,1).font.bold = true

    cells(2,1).value =  1
    cells(3,1).value =  1

    range(cells(4,1), cells(9,1)).formulaR1C1 = "= R[-2]C + R[-1]C"

    cells(1,3).value     = "miscellaneous"
    cells(1,3).font.bold = true
    cells(2,3).value =  18
    cells(3,3).value =  8
    cells(4,3).value = "foo"
    cells(5,3).value =  17
    cells(6,3).value = "bar"
    cells(7,3).value = "baz"
    cells(8,3).value =  8
    cells(9,3).value =  17

end sub ' }
