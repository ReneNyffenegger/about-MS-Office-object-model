option explicit

sub main() ' {


    dim ary_2d as variant

    dim referenceStyleOrig as long : referenceStyleOrig = application.referenceStyle

  '
  ' R1C1 really makes things a lot easier:
  '
    application.referenceStyle = xlR1C1

  '
  ' Create a 4x3 two-dimensional array and initialize its values to 42
  '
  ' Beware: if must be immediately followed by paranthesis.
  '
    ary_2d = [ if( isError(r1c1:r4c3), 42, 42 ) ]

    application.referenceStyle = referenceStyleOrig

    debug.print "Dimensions of created array are: "          & _
       lBound(ary_2d, 1) & " to " & uBound(ary_2d, 1) & ", " & _
       lBound(ary_2d, 2) & " to " & uBound(ary_2d, 2)

    debug.print "ary(2,3) = " & ary_2d(2,3)

end sub ' }
