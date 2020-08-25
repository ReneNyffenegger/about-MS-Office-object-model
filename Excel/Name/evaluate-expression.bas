option explicit

sub main() ' {

    activeSheet.name = "expressions"

    cells(2,2) = "2+3+4+5"
    cells(3,2) = "22/3"
    cells(4,2) = "sin(3.1)"

    activeSheet.names.add _
       name          := "evalExprToTheRight"             , _
       refersToR1C1  := "=evaluate( expressions!rc[-1])"

    cells(2,3).formulaR1C1 = "=evalExprToTheRight"
    cells(2,3).autoFill                               _
       destination := range(cells(2,3), cells(4,3)) , _
       type        := xlFillCopy

end sub ' }
