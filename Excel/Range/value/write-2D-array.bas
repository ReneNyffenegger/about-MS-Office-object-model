option explicit

sub main() ' {

    dim array2D as variant
  '
  ' Use application.evaluate to create a 2D array:
  '
    array2D = evaluate(                                                    _
                  "{ """"         , ""Val 1"", ""Val 2"", ""Val 3""  ; " & _
                  "  ""Row one""  ,      17  ,       29 ,        18  ; " & _
                  "  ""Row two""  ,       4  ,       13 ,        12  ; " & _
                  "  ""Row three"",      16  ,       25 ,         7  ; " & _
                  "  ""Row four"" ,      22  ,        9 ,        14  } " )


    cells(3,2).resize(uBound(array2D, 1), ubound(array2D, 2)).value = array2D
    cells(3,2).resize(1                 , ubound(array2D, 2)).entireColumn.autoFit

end sub ' }
