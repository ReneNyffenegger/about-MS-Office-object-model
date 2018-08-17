option explicit

sub main() ' {

    dim val as double

    dim wsf as worksheetFunction : set wsf = application.worksheetFunction
    dim rng as range             : set rng = range("a1:b5")

    fillTestData

    val = wsf.min    (rng) : debug.print "min = " & val
    val = wsf.sum    (rng) : debug.print "sum = " & val
    val = wsf.average(rng) : debug.print "avg = " & val

end sub ' }

sub fillTestData() ' {

    cells(1, 1) =  9: cells(1, 2) =  5
    cells(2, 1) =  2: cells(2, 2) =  8
    cells(3, 1) =  4: cells(3, 2) =  7
    cells(4, 1) =  5: cells(4, 2) =  5
    cells(5, 1) =  3: cells(5, 2) =  4

end sub ' }
