'
'  Inspired by (and mostly copied from) https://stackoverflow.com/a/43619770/180275
'
option explicit

sub main() ' {

    cells(1, 1) = "Number" : cells(1,2) =  42
    cells(2, 1) = "Text"   : cells(2,2) = "Hello world"

    activeSheet.usedRange.columns.autoFit

    saveRangeAsImage _
       range(cells(1,1), cells(2,2)), _
       environ("temp") & "\exportedRange.png"

end sub ' }

sub saveRangeAsImage(rng as range, fileName as string) ' {

    rng.copyPicture               _
        appearance := xlScreen  , _
        format     := xlPicture

    dim crt as chartObject
    set crt = rng.parent.chartObjects.add(0, 0, rng.width, rng.height)

  '
  ' These doEvents seem necessary to give some time to Excel
  ' to complete some operation
  '
    doEvents
    doEvents
    doEvents

    crt.chart.paste
    crt.chart.export filename := fileName, filtername := "png"

    crt.delete

end sub ' }
