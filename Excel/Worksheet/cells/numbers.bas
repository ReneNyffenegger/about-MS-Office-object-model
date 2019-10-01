option explicit

sub main() ' {

    dim allCells as range
    set allcells = activeSheet.cells()

    debug.print "The maximum number of rows in a worksheet is: " & allCells.rows.   count
    debug.print "The maximum number of cols in a worksheet is: " & allCells.columns.count



  ' Get maximum number of cells in an Excel Worksheet
  '
  '    Use countLarge because count throws
  '    a runtime error 6 (Overflow):
    dim maxCntOfCellsOnSheet as longLong
    maxCntOfCellsOnSheet = allCells.countLarge
    debug.print "This corresponds to  " & maxCntOfCellsOnSheet & " cells"

    debug.print "The address of 'all cells' is " & allCells.address

end sub ' }
'
' The maximum number of rows in a worksheet is: 1048576
' The maximum number of cols in a worksheet is: 16384
' This corresponds to  17179869184 cells
' The address of 'all cells' is $1:$1048576
