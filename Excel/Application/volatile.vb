option explicit

global nofVolatileCalls    as long
global nofNonVolatileCalls as long

function volatileFunction() as long ' {

    application.volatile

    nofVolatileCalls = nofVolatileCalls + 1
    volatileFunction = nofVolatileCalls

end function ' }

function nonVolatileFunction() as long ' {

    nofNonVolatileCalls = nofNonVolatileCalls + 1
    nonVolatileFunction = nofNonVolatileCalls

end function ' }

sub main() ' {

  ' Create worksheet with these functions

    dim ws as worksheet
    set ws = activeSheet

    ws.usedRange.clearContents
    ws.usedRange.clearFormats

    ws.cells(2,2) = "Volatile function:"
    ws.cells(3,2) = "Non-Volatile function:"

    ws.cells(2,3).formula = "= volatileFunction()"
    ws.cells(3,3).formula = "= NonvolatileFunction()"

    ws.usedRange.columns.autoFit

end sub ' }
