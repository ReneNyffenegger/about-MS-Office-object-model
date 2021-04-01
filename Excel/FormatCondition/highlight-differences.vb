option explicit

sub main() ' {

    cells(1,1) = "foo" : cells(1,2) = "bar": cells(1,3) = "baz" : cells(1,4) = 42 : cells(1,5) = 18
    cells(2,1) = "foo" : cells(2,2) = "bar": cells(2,3) = "baz" : cells(2,4) = 99 : cells(2,5) = 18

    dim fcs as formatConditions
    set fcs = range(cells(2,1), cells(2,5)).formatconditions

 '
 '  German Excel: use "=Z(-1)S" (https://stackoverflow.com/a/48539578/180275)
 '                for formula.
 '
    dim fc as formatCondition
    set fc = fcs.add(xlCellValue, xlNotEqual, "=R[-1]C")

    fc.interior.color = rgb(255, 170, 170)

end sub ' }
