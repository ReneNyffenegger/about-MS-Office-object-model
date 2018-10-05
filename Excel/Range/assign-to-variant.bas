option explicit

sub main() ' {

    testData

    dim ary2d as variant

    ary2d = range(cells(1,1), cells(5,4))

    dim r as long
    dim c as long
    for r = 1 to 5: for c = 1 to 4
        debug.print "ary2d(" & r & ", " & c & ") = "  ary2d(r, c)
    next c: next r

end sub ' }

sub testData() ' {

    activeSheet.usedRange.clearContents

    dim r as long
    dim c as long
    for r = 1 to 5: for c = 1 to 4
        cells(r, c) = (r mod 3 + c mod 2) * (c mod 2 + 1)
    next c: next r

end sub ' }
