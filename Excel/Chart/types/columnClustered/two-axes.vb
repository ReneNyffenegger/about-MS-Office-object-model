option explicit

sub main() ' {

    fillData

    dim shap as shape

    cells(1,1).currentRegion.select

    set shap = activeSheet.shapes.addChart2(201, xlColumnClustered)
    dim chrt as chart : set chrt = shap.chart

    chrt.fullSeriesCollection(2).axisGroup = 2
    chrt.fullSeriesCollection(3).axisGroup = 2
    chrt.fullSeriesCollection(4).axisGroup = 2
    chrt.fullSeriesCollection(5).axisGroup = 2

    shap.left   =  300
    shap.width  =  600
    shap.top    =  200
    shap.height =  300

    chrt.fullSeriesCollection(1).name = "Month"
    chrt.fullSeriesCollection(2).name = "Val 1"
    chrt.fullSeriesCollection(3).name = "Val 2"
    chrt.fullSeriesCollection(4).name = "Val 3"
    chrt.fullSeriesCollection(5).name = "Val 4"


    chrt.chartGroups(1).gapWidth =  33
    chrt.chartGroups(2).gapWidth = 200
    chrt.chartGroups(2).overlap  = -19

    dim lgnd as legend : set lgnd = chrt.legend
  '
  ' Note, the legendKey attribute is not recorded in the Macro Recorder!
  '                   '
    lgnd.legendEntries(1).legendKey.format.fill.foreColor.rgb = rgb(180, 180, 180)
    lgnd.legendEntries(2).legendKey.format.fill.foreColor.rgb = rgb(255,  50,  50)
    lgnd.legendEntries(3).legendKey.format.fill.foreColor.rgb = rgb( 50, 100, 255)
    lgnd.legendEntries(4).legendKey.format.fill.foreColor.rgb = rgb(255, 195,  15)
    lgnd.legendEntries(5).legendKey.format.fill.foreColor.rgb = rgb( 30, 200,  40)

    lgnd.width = 320
    lgnd.left  = 145

    chrt.chartTitle.text = "Values in the Year 2021"

    chrt.axes(xlCategory).tickLabels.orientation = xlUpward

    chrt.export fileName := environ("temp") & "\two-axes.png", filterName := "png"

end sub ' }


sub fillData() ' {

    dim rng as range : set rng=range(cells(1,1), cells(1,6))

    rng.value = array("2021-01", 13380, 52, 45, 44, 12, 43) : set rng=rng.offset(1)
    rng.value = array("2021-02", 11825, 49, 43, 28, 23, 46) : set rng=rng.offset(1)
    rng.value = array("2021-03", 12778, 46, 34, 36, 24, 52) : set rng=rng.offset(1)
    rng.value = array("2021-04", 13549, 46, 47, 50, 28, 50) : set rng=rng.offset(1)
    rng.value = array("2021-05", 13888, 49, 50, 59, 21, 46) : set rng=rng.offset(1)
    rng.value = array("2021-06", 12536, 47, 33, 44, 29, 42) : set rng=rng.offset(1)
    rng.value = array("2021-07", 14045, 50, 38, 45, 18, 47) : set rng=rng.offset(1)
    rng.value = array("2021-08", 14219, 47, 26, 62, 17, 53) : set rng=rng.offset(1)
    rng.value = array("2021-09", 13268, 41, 39, 61, 27, 46) : set rng=rng.offset(1)
    rng.value = array("2021-10", 13298, 57, 42, 43, 22, 34) : set rng=rng.offset(1)
    rng.value = array("2021-11", 14040, 44, 28, 53, 16, 47) : set rng=rng.offset(1)
    rng.value = array("2021-12", 12263, 64, 35, 41, 26, 46) : set rng=rng.offset(1)

end sub ' }