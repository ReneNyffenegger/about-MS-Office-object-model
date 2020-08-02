option explicit

sub main() ' {

    dim dataAll as range, dataValues as range, dataCategoryNames as range
    set dataAll = testData

    set dataValues        = application.intersect(dataAll, dataAll.offset(0,1))
    set dataCategoryNames = application.intersect(dataAll, dataAll.offset(1,0)).resize(columnSize := 1)

  '
  '     Although this example is going to create a combo chart,
  '     it starts with a clustered column chart.
  '
    dim sh as shape
    set sh = activeSheet.shapes.addChart2(201, xlColumnClustered)

    dim cht as chart
    set cht = sh.chart

    with cht ' {

      .setSourceData source := dataValues

      '
      '   Add/define clustered column series for first set
      '   of values.
      '
      .fullSeriesCollection(1).chartType = xlColumnClustered
      .fullSeriesCollection(1).axisGroup = 1
      
      '
      '   Add/define line series for second set of values.
      '
      '   Because the first and second series have different
      '   chart type, the chart becomes a combo chart.
      '
      .fullSeriesCollection(2).chartType = xlLine
      .fullSeriesCollection(2).axisGroup = 1
      .fullSeriesCollection(2).axisGroup = 2

       with .axes(xlcategory)
              .hastitle          =  true
              .axisTitle.caption = "Development of values"
              .categoryNames     =  dataCategoryNames
       end with 

      .axes(xlValue, xlSecondary).minimumScale = 14

      .hasTitle        =  true
      .chartTitle.text = "Combo Chart Example"

      .export                                                       _
         fileName   :=  activeWorkbook.path & "\img\line-and-column.png", _
         filterName := "png"

    end with ' }

end sub ' }

function testData() as range ' {

    dim r as long
    r = 2

    with activeSheet ' {
      '
      '  Insert test data.
      '
        .range(.cells(r, 2), .cells(r, 4)).font.bold = true

        .range(.cells(r, 2), .cells(r, 4)).value = array( "Year", "val one", "val two") : r=r+1
                                                        '  ---- ,  ------- ,  --------
        .range(.cells(r, 2), .cells(r, 4)).value = array(  2015 ,      115 ,     15.7 ) : r=r+1
        .range(.cells(r, 2), .cells(r, 4)).value = array(  2016 ,      117 ,     16.5 ) : r=r+1
        .range(.cells(r, 2), .cells(r, 4)).value = array(  2017 ,      112 ,     24.9 ) : r=r+1
        .range(.cells(r, 2), .cells(r, 4)).value = array(  2018 ,      109 ,     22.8 ) : r=r+1
        .range(.cells(r, 2), .cells(r, 4)).value = array(  2019 ,      116 ,     16.1 ) : r=r+1
        .range(.cells(r, 2), .cells(r, 4)).value = array(  2020 ,      118 ,     18.3 )

         set testData = range(cells(2,2), cells(r,4))

    end with ' }

end function ' }
