option explicit

sub main() ' {

    createTestData

    dim sh as shape
    set sh = activeSheet.shapes.addChart2(201, xlColumnClustered)

    sh.top    =                     cells( 2, 6).top
    sh.left   =                     cells( 2, 6).left
    sh.height = cells(15,12).top  - cells( 2, 6).top
    sh.width  = cells(15,12).left - cells( 2, 6).left

    dim ch as chart
    set ch = sh.chart

    ch.setSourceData source := range(cells(2,2), cells(5,4))
    ch.setElement msoElementDataTableWithLegendKeys

end sub ' }

sub createTestData() ' {

    dim r as long
    r = 2
    range(cells(r,2), cells(r,4)).value = array( null, 2019, 2018) : r=r+1
    range(cells(r,2), cells(r,4)).value = array("foo",    5,    6) : r=r+1
    range(cells(r,2), cells(r,4)).value = array("bar",    4,    3) : r=r+1
    range(cells(r,2), cells(r,4)).value = array("baz",    7,    5)

end sub ' }
