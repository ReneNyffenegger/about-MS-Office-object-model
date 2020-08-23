option explicit

sub main() ' {

    activeSheet.names.add "romanNumbers", refersTo := "={""I"", ""II"", ""III"", ""IV"", ""V"", ""VI"", ""VII"", ""VIII"", ""IX"", ""X"", ""XI"", ""XII""}"

    cells(2,2) =  7
    cells(3,2) =  5
    cells(4,2) = 11
    cells(5,2) =  3
    cells(6,2) =  6

    range(cells(2,3), cells(6,3)).formulaR1C1 = "= index(romanNumbers, RC[-1])"

end sub ' }
