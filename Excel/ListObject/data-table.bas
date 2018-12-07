option explicit

sub main() ' {

  '
  ' Create some t est data
  '
    testData

  '
  ' Create a data table from the test data:
  '
    dim dataTable as listObject
    set dataTable = activeSheet.listObjects.add(xlSrcRange, range(cells(3,4), cells(12,6)))

    dataTable.name = "datTbl"

  '
  ' Add a totals row at the bottom of the table:
  '
    dataTable.showTotals = true

  '
  ' Show maximum number of second column (column name of which
  ' is colTwo)
  '
    dataTable.listColumns("colTwo").totalsCalculation = xlTotalsCalculationMax


end sub ' }

sub testData() ' {

    cells( 3, 4) = "colOne" : cells( 3, 5) = "colTwo" : cells( 3, 6) = "colThree"
    cells( 4, 4) = "bar"    : cells( 4, 5) =      15  : cells( 4, 6) =        34
    cells( 5, 4) = "foo"    : cells( 5, 5) =      21  : cells( 5, 6) =        30
    cells( 6, 4) = "baz"    : cells( 6, 5) =      20  : cells( 6, 6) =        35
    cells( 7, 4) = "bar"    : cells( 7, 5) =      18  : cells( 7, 6) =        29
    cells( 8, 4) = "foo"    : cells( 8, 5) =      16  : cells( 8, 6) =        31
    cells( 9, 4) = "foo"    : cells( 9, 5) =      21  : cells( 9, 6) =        36
    cells(10, 4) = "bar"    : cells(10, 5) =      18  : cells(10, 6) =        34
    cells(11, 4) = "baz"    : cells(11, 5) =      19  : cells(11, 6) =        32
    cells(12, 4) = "foo"    : cells(12, 5) =      17  : cells(12, 6) =        31

end sub ' }
