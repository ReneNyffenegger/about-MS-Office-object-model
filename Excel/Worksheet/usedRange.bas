option explicit

sub main() ' {

    dim sh as workSheet
    set sh = activeSheet

    debug.print sh.usedRange.address(referenceStyle := xlR1C1) ' R1C1

    cells(7, 5) = "foo"
    debug.print sh.usedRange.address(referenceStyle := xlR1C1) ' R7C5

    cells(3, 8) = "bar"
    debug.print sh.usedRange.address(referenceStyle := xlR1C1) ' R3C5:R7C8

    sh.usedRange.clearContents
    debug.print sh.usedRange.address(referenceStyle := xlR1C1) ' R1C1

end sub ' }
