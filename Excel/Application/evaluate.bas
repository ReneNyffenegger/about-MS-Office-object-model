option explicit

sub main() ' {

    fillTestData

    evaluateFuncOnTestdata "sum"
    evaluateFuncOnTestdata "min"
    evaluateFuncOnTestdata "average"

end sub ' }

sub fillTestData() ' {

    cells(1, 1) =  9: cells(1, 2) =  5
    cells(2, 1) =  2: cells(2, 2) =  8
    cells(3, 1) =  4: cells(3, 2) =  7
    cells(4, 1) =  5: cells(4, 2) =  5
    cells(5, 1) =  3: cells(5, 2) =  4

end sub ' }

sub evaluateFuncOnTestdata(funcName as string) ' {

    debug.print funcName & " = " & application.evaluate(funcName & "(a1:b5)")

end sub ' }
