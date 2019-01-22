option explicit

sub testData() ' {

    cells (1, 1) = "foo": cells (1, 2) = 4: cells(1, 3).formulaR1C1 = "=repeatString(r[0]c[-2], r[0]c[-1])"
    cells (2, 1) = "bar": cells (2, 2) = 1: cells(2, 3).formulaR1C1 = "=repeatString(r[0]c[-2], r[0]c[-1])"
    cells (3, 1) = "baz": cells (3, 2) = 3: cells(3, 3).formulaR1C1 = "=repeatString(r[0]c[-2], r[0]c[-1])"

end sub ' }

function repeatString(cellText as range, cellTimes as range) as string ' {

    dim i as long
    repeatString = cellText
    for i = 2 to cellTimes
        repeatString = repeatString & " " & cellText
    next i

end function ' }
