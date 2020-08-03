option explicit

sub main() ' {

    dim rngData as range
    set rngData = range( cells(1,1), cells(9,1) )
    rngData.value = worksheetFunction.transpose(array("Data", 2.1 , 3.8 , 2.9 , -1.4 , -2.0 , 1.1 , 2.8, -0.4 ))

  '
  ' Apparently, a waterfall chart does not allow to set the
  ' initial source range with .setSourceData, it seems that
  ' the source data needs to be selected when the chart is
  ' created:
  '
    rngData.select

    dim shp as shape
    set shp = activeSheet.shapes.addChart2(395, xlWaterfall)

    dim cht   as  chart
    set cht = shp.chart

    with cht ' {

      .hasTitle        =  true
      .chartTitle.text = "Simple Waterfall Chart"

      .export                                                    _
         fileName   :=  activeWorkbook.path & "\img\simple.png", _
         filterName := "png"

    end with ' }

end sub ' }
