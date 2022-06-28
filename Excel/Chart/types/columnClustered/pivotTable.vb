option explicit

sub main() ' {

    dim testData as range

    set testData = createTestData

    createChart testData

end sub ' }

sub createChart(rng as range) ' {

    dim pc as pivotCache

    set pc = activeWorkbook.pivotCaches.create(sourceType:=xlDatabase, sourceData:= rng)

    dim pt as pivotTable
    set pt = pc.createPivotTable(tableDestination := rng.parent.cells(3, 4))


    pt.RowGrand = false

    cells(3, 5).select
    dim sh   as shape
    set sh = activeSheet.shapes.addChart2(201, xlColumnClustered)

    dim ch   as chart
    set ch = sh.chart

    sh.select

    ch.setSourceData source := range(pt.tableRange1.address)

    pt.pivotFields("Item").orientation = xlColumnField

    dim  fl as pivotField
    set  fl = pt.addDataField(pt.pivotFields("Value"), "Value-Sum", xlSum)
    with fl
        .caption  = "Count of Value"
        .function = xlCount
    end with

    pt.pivotFields("Value").orientation = xlRowField

    ch.chartColor = 12

  '
  ' Place shape
  '
    dim rngShape as range
    with rng.parent : set rngShape = .range(.cells(3, 9), .cells(20, 19)) : end with

    sh.left   = rngShape.left
    sh.top    = rngShape.top
    sh.width  = rngShape.width
    sh.height = rngShape.height
  '
  ' Format legend
  '
    dim lg   as legend
    set lg = ch.legend
  '
  ' Delete legent if not needed
  ' lg.delete
  '
    lg.includeInLayout = false
    lg.top  =  10
    lg.left = 410

    lg.format.line.foreColor.rgb = rgb(100, 50, 160)
  '
  ' Hide buttons
  '
    ch.showAllFieldButtons    = false
  '
  ' Hide different types of field buttons
  ' individually:
  '
  ' ch.showValueFieldButtons  = false
  ' ch.showAxisFieldButtons   = false
  ' ch.showLegendFieldButtons = false

  '
  ' Delete (horizontal) chart lines
  '
    ch.axes(xlValue).majorGridLines.delete

  '
  ' Save chart as image
  ' https://renenyffenegger.ch/notes/Microsoft/Office/Excel/Object-Model/Range/copyPicture/save-range-as-image
  '

    ch.export fileName := environ("temp") & "\bar-chart.png", filterName := "png"

end sub ' }

function createTestData() as range ' {

   dim ws as worksheet
   set ws = worksheets.add

   ws.activate

   dim row_ as long
   row_ = 1
   ws.cells(row_, 1) = "Item"
   ws.cells(row_, 2) = "Value"

   for row_ = 2 to 1001 ' {

       dim rnd_ as double
       dim val_ as long
       dim itm_ as string

       rnd_ = rnd

       if     rnd_ < 0.3 then

              itm_ = "foo"
              val_ =  50 + rnd * 12

       elseif rnd_ < 0.8 then

              rnd_ =  rnd
              itm_ = "bar"
              val_ =  49 + rnd_ * rnd_ * 16

       else
              itm_ = "baz"
              rnd_ =  rnd
              val_ =  52 + (1-rnd_) * (1-rnd_) * 8

       end if

       ws.cells(row_, 1) = itm_
       ws.cells(row_, 2) = val_

   next row_ ' }

   set createTestData = ws.range(ws.cells(1,1), ws.cells(row_-1, 2))

end function ' }
