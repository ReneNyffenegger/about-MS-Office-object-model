option explicit

sub main() ' {

    dim rng as range

    set rng = cells(7, 3)

    debug.print (rng.address)             ' $C$7

    debug.print (rng.address(   _
       rowAbsolute    := false, _
       columnAbsolute := false  _
    ))                                    ' C7

    debug.print (rng.address(   _
       referenceStyle := xlR1C1 _
    ))                                    ' R7C3


    debug.print (rng.address(        _
       rowAbsolute    := false     , _
       columnAbsolute := false     , _
       referenceStyle := xlR1C1    , _
       relativeTo     := cells(5, 8) _
   ))                                    ' R[2]C[-5]


end sub ' }
