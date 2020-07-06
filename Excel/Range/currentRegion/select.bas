option explicit

sub main() ' {

    testData

    dim region as range

    set region = cells(3,4).currentRegion
    region.interior.color = rgb(255, 245, 235)
    region.columns.autoFit

end sub ' }

sub testData() ' {

    cells(2, 2) = 42
    cells(2, 3) = "x"
    cells(3, 3) =  8
    cells(3, 5) = "?!"
    cells(3, 4) = "foo"
    cells(4, 2) = 9.31

end sub ' }
