option explicit

const tl_r =  3
const tl_c =  2
const w    =  4
const h    =  9

sub main() ' {

  '
  '    Cannot set screenUpdating to false because the
  '    application.goto (below) would not work
  '    anymore.
  '
  ' application.screenUpdating    = false

    activeWindow.displayGridlines = false

    dim wsh as worksheet
    set wsh =  worksheets.add

    dim r as long : r=0
    dim c as long : c=0

    createChart wsh, r, c, xl3DArea                  , "xl3DArea"
    createChart wsh, r, c, xl3DAreaStacked           , "xl3DAreaStacked"
    createChart wsh, r, c, xl3DAreaStacked100        , "xl3DAreaStacked100"
    createChart wsh, r, c, xl3DBarClustered          , "xl3DBarClustered"
    createChart wsh, r, c, xl3DBarStacked            , "xl3DBarStacked"
    createChart wsh, r, c, xl3DBarStacked100         , "xl3DBarStacked100"
    createChart wsh, r, c, xl3DColumn                , "xl3DColumn"
    createChart wsh, r, c, xl3DColumnClustered       , "xl3DColumnClustered"
    createChart wsh, r, c, xl3DColumnStacked         , "xl3DColumnStacked"
    createChart wsh, r, c, xl3DColumnStacked100      , "xl3DColumnStacked100"
    createChart wsh, r, c, xl3DLine                  , "xl3DLine"
    createChart wsh, r, c, xl3DPie                   , "xl3DPie"
    createChart wsh, r, c, xl3DPieExploded           , "xl3DPieExploded"
    createChart wsh, r, c, xlArea                    , "xlArea"
    createChart wsh, r, c, xlAreaStacked             , "xlAreaStacked"
    createChart wsh, r, c, xlAreaStacked100          , "xlAreaStacked100"
    createChart wsh, r, c, xlBarClustered            , "xlBarClustered"
    createChart wsh, r, c, xlBarOfPie                , "xlBarOfPie"
    createChart wsh, r, c, xlBarStacked              , "xlBarStacked"
    createChart wsh, r, c, xlBarStacked100           , "xlBarStacked100"
    createChart wsh, r, c, xlBubble                  , "xlBubble"
    createChart wsh, r, c, xlBubble3DEffect          , "xlBubble3DEffect"
    createChart wsh, r, c, xlColumnClustered         , "xlColumnClustered"
    createChart wsh, r, c, xlColumnStacked           , "xlColumnStacked"
    createChart wsh, r, c, xlColumnStacked100        , "xlColumnStacked100"
    createChart wsh, r, c, xlConeBarClustered        , "xlConeBarClustered"
    createChart wsh, r, c, xlConeBarStacked          , "xlConeBarStacked"
    createChart wsh, r, c, xlConeBarStacked100       , "xlConeBarStacked100"
    createChart wsh, r, c, xlConeCol                 , "xlConeCol"
    createChart wsh, r, c, xlConeColClustered        , "xlConeColClustered"
    createChart wsh, r, c, xlConeColStacked          , "xlConeColStacked"
    createChart wsh, r, c, xlConeColStacked100       , "xlConeColStacked100"
    createChart wsh, r, c, xlCylinderBarClustered    , "xlCylinderBarClustered"
    createChart wsh, r, c, xlCylinderBarStacked      , "xlCylinderBarStacked"
    createChart wsh, r, c, xlCylinderBarStacked100   , "xlCylinderBarStacked100"
    createChart wsh, r, c, xlCylinderCol             , "xlCylinderCol"
    createChart wsh, r, c, xlCylinderColClustered    , "xlCylinderColClustered"
    createChart wsh, r, c, xlCylinderColStacked      , "xlCylinderColStacked"
    createChart wsh, r, c, xlCylinderColStacked100   , "xlCylinderColStacked100"
    createChart wsh, r, c, xlDoughnut                , "xlDoughnut"
    createChart wsh, r, c, xlDoughnutExploded        , "xlDoughnutExploded"
    createChart wsh, r, c, xlFunnel                  , "xlFunnel"
    createChart wsh, r, c, xlLine                    , "xlLine"
    createChart wsh, r, c, xlLineMarkers             , "xlLineMarkers"
    createChart wsh, r, c, xlLineMarkersStacked      , "xlLineMarkersStacked"
    createChart wsh, r, c, xlLineMarkersStacked100   , "xlLineMarkersStacked100"
    createChart wsh, r, c, xlLineStacked             , "xlLineStacked"
    createChart wsh, r, c, xlLineStacked100          , "xlLineStacked100"
    createChart wsh, r, c, xlPie                     , "xlPie"
    createChart wsh, r, c, xlPieExploded             , "xlPieExploded"
    createChart wsh, r, c, xlPieOfPie                , "xlPieOfPie"
    createChart wsh, r, c, xlPyramidBarClustered     , "xlPyramidBarClustered"
    createChart wsh, r, c, xlPyramidBarStacked       , "xlPyramidBarStacked"
    createChart wsh, r, c, xlPyramidBarStacked100    , "xlPyramidBarStacked100"
    createChart wsh, r, c, xlPyramidCol              , "xlPyramidCol"
    createChart wsh, r, c, xlPyramidColClustered     , "xlPyramidColClustered"
    createChart wsh, r, c, xlPyramidColStacked       , "xlPyramidColStacked"
    createChart wsh, r, c, xlPyramidColStacked100    , "xlPyramidColStacked100"
    createChart wsh, r, c, xlRadar                   , "xlRadar"
    createChart wsh, r, c, xlRadarFilled             , "xlRadarFilled"
    createChart wsh, r, c, xlRadarMarkers            , "xlRadarMarkers"
    createChart wsh, r, c, xlRegionMap               , "xlRegionMap"
    createChart wsh, r, c, xlStockHLC                , "xlStockHLC"
    createChart wsh, r, c, xlStockOHLC               , "xlStockOHLC"
    createChart wsh, r, c, xlStockVHLC               , "xlStockVHLC"
    createChart wsh, r, c, xlStockVOHLC              , "xlStockVOHLC"
    createChart wsh, r, c, xlSurface                 , "xlSurface"
    createChart wsh, r, c, xlSurfaceTopView          , "xlSurfaceTopView"
    createChart wsh, r, c, xlSurfaceTopViewWireframe , "xlSurfaceTopViewWireframe"
    createChart wsh, r, c, xlSurfaceWireframe        , "xlSurfaceWireframe"
    createChart wsh, r, c, xlXYScatter               , "xlXYScatter"
    createChart wsh, r, c, xlXYScatterLines          , "xlXYScatterLines"
    createChart wsh, r, c, xlXYScatterLinesNoMarkers , "xlXYScatterLinesNoMarkers"
    createChart wsh, r, c, xlXYScatterSmooth         , "xlXYScatterSmooth"
    createChart wsh, r, c, xlXYScatterSmoothNoMarkers, "xlXYScatterSmoothNoMarkers"
    createChart wsh, r, c, xlCombination             , "xlCombination"

  ' application.screenUpdating = true

end sub ' }

sub createChart(wsh as worksheet, byRef r as long, byRef c as long, chtType as xlChartType, chtTypeTxt as string) ' {

 on error goto err_

    dim shp as shape
    dim cht as chart

    set shp    = wsh.shapes.addChart2(xlChartType := chtType)

    dim rng as range
    set rng    = wsh.range(wsh.cells(1+tl_r + r*(h+1)       , 1+tl_c + c*(w+1)      ), _
                           wsh.cells(1+tl_r + r*(h+1)+h - 1 , 1+tl_c + c*(w+1)+w - 1)  _
                          )

    shp.left   = rng.left
    shp.top    = rng.top
    shp.width  = rng.width
    shp.height = rng.height

    set cht = shp.chart

    dim ser_1 as series
    dim ser_2 as series
    dim ser_3 as series

    with cht.seriesCollection ' {

         while .count > 0 ' {
               .item(1).delete
         wend ' }

         set ser_1 = .newSeries
         set ser_2 = .newSeries
         set ser_3 = .newSeries

    end with ' }

    ser_1.values = array(145,  89, 120) : ser_1.xValues = array("foo", "bar", "baz") : ser_1.name = "one"
    ser_2.values = array( 30,  71,  34) : ser_2.xValues = array("foo", "bar", "baz") : ser_2.name = "two"
    ser_3.values = array(418, 505, 654) : ser_3.xValues = array("foo", "bar", "baz") : ser_3.name = "three"
  '
  ' I don't really understand why the call of chartWizard is necessary, but without it,
  ' all charts but the first are not drawn.
  '
    cht.chartWizard
  '
  ' Trying to set change the axis group of the third group.
  ' This is not possible for all chart types, hence the
  ' special error handling here.
  '
    on error resume next
    ser_3.axisGroup = xlSecondary
    if err.number = 1004 then ' Parameter not valid
       debug.print "Axis group cannot be set for " & chtTypeTxt
    end if
    on error goto err_

    cht.hasTitle = true
    cht.chartTitle.text = chtTypeTxt

    if c >= 3 then ' {
       c = 0
       r = r + 1
    else
       c = c + 1
    end if ' }

  '
  ' The chart needs to be in the visible range when exported, otherwise
  ' empty files will produced.
  ' Therefore: go to the chart:
  '
    application.goto rng
    cht.export _
        fileName   :=  wsh.parent.path & "\img\" & chtTypeTxt & ".png", _
        filterName := "png"

    exit sub

err_:

    debug.print err.description & " (" & err.number & ") for " & chtTypeTxt

end sub ' }
