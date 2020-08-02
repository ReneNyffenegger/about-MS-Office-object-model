option explicit

const nofPoints = 25
const pi        = 3.141592

sub main() ' {

    with activeSheet ' {

        .cells( 1, 1) = "t"
        .cells( 1, 2) = "x"
        .cells( 1, 3) = "y"

        .cells(18, 5) = "a" : .cells(18, 6) = 3
        .cells(19, 5) = "b" : .cells(19, 6) = 4

        .range( .cells(2,1), .cells(nofPoints+1,1) ).formulaR1C1 = "= 2 * " & pi & " * (row()-2) / " & (nofPoints-1)

        .range( .cells(2,2), .cells(nofPoints+1,2) ).formulaR1C1 = "= sin( rc[-1] * r18c6 )"
        .range( .cells(2,3), .cells(nofPoints+1,3) ).formulaR1C1 = "= sin( rc[-2] * r19c6 )"

         dim shp as shape
         set shp = .shapes.addChart2(xlChartType := xlXYScatterSmoothNoMarkers)

         dim cht as chart
         set cht =  shp.chart

         cht.setSourceData source := .range(.cells(2,2), .cells(nofPoints+1,3))

         with .range(.cells(2,4), cells(16,11) ) ' {

             shp.width  = .width
             shp.height = .height
             shp.top    = .top
             shp.left   = .left

         end with ' }

         cht.hasTitle        =  true
         cht.chartTitle.text = "Lissajous"

         cht.legend.delete

         cht.export _
             fileName   :=  activeWorkbook.path & "\img\lissajous.png", _
             filterName := "png"

    end with ' }

end sub ' }
