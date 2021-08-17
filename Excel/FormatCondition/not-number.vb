option explicit

sub main() ' {

   dim rng as range
   set rng = range(cells(4,3), cells(7, 7))
   rng.borderAround xlContinuous, xlMedium, color := rgb(100, 150, 200)

   highlightNonNumbersInRange rng


   cells(4,2) = "foo"
   cells(4,3) = "bar"
   cells(4,4) = "baz"

   cells(3,5) =  42
   cells(4,5) =  99
   cells(5,5) =  13.12

   cells(7,4) = "hello"
   cells(8,4) = "world"

   cells(6,6) = "xyz"
   cells(7,6) = 12345678

   cells(10,10).select

end sub ' }


sub highlightNonNumbersInRange(rng as range) ' {

    dim formula as string
'   formula = "=not(isNumber(" & rng.cells(1).address(rowAbsolute := false, columnAbsolute := false) & "))"
    formula = "=not(isNumber(r[0]c[0]))"

    dim fc as formatCondition

    set fc = rng.formatConditions.add(xlExpression, formula1 := formula)

    fc.font.color = rgb(230, 20, 40)

end sub ' }
