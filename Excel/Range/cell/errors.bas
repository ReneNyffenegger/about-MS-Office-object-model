option explicit

sub main() ' {
    createCellsWithErrors

    checkIfCellContainsError  1, 2
    checkIfCellContainsError  2, 2
    checkIfCellContainsError  3, 2
    checkIfCellContainsError  4, 2
    checkIfCellContainsError  5, 2
    checkIfCellContainsError  6, 2
    checkIfCellContainsError  7, 2

    checkIfCellContainsError  1, 3
end sub ' }


sub createCellsWithErrors() ' {

    cells(1, 1) = "xlErrDiv0"
    cells(1, 3) =  42
    cells(1, 4) =   0
    cells(1, 2).formulaR1C1 = "= RC[1] / RC[2]"

    cells(2, 1) = "xlErrNA"
    cells(2, 3) = "foo"
    cells(2, 4) = "bar"
    cells(2, 5) = "baz"
    cells(2, 6) = "n.a."
    cells(2, 2).formulaR1C1 = "=match(RC[4],RC[1]:RC[3], 0)"

    cells(3, 1) = "xlErrName"
    cells(3, 2).formulaR1C1 = "= inexistingFunction()"

    cells(4, 1) = "xlErrNull"
    cells(4, 2).formulaR1C1 = "= sum(RC[1]:RC[2] RC[3]:RC[4])" ' Use intersect operator (space) on two ranges that have no cell in common

    cells(5, 1) = "xlErrNum"
    cells(5, 3).formulaR1C1 = "= 17.9769313486231 * 1e307" ' Close to the maximum double precision floating-point value
    cells(5, 2).formulaR1C1 = "= 17.9769313486232 * 1e307" ' Over the maximum double precision floating-point value

    cells(6, 1) = "xlErrRef"
    cells(6, 9) = "some value that is going to be deleted"
    cells(6, 2).formulaR1C1 = "= RC9"
    columns(9).delete

    cells(7, 1) = "xlErrValue"
    cells(7, 2).formulaR1C1 = "= ""foo"" + ""bar""" ' cannot add two strings...

    range(columns(1), columns(5)).autoFit
    cells(9, 9).select 

end sub ' }

sub checkIfCellContainsError(r as long, c as integer) ' {

    dim cellVal as variant
    cellVal = cells(r, c).value

    if varType(cellVal) = vbError then
       dim errText as string

       select case cellVal
              case cvErr(xlErrDiv0 ): errText = "xlErrDiv0"
              case cvErr(xlErrNA   ): errText = "xlErrNA"
              case cvErr(xlErrName ): errText = "xlErrName"
              case cvErr(xlErrNull ): errText = "xlErrNull"
              case cvErr(xlErrNum  ): errText = "xlErrNum"
              case cvErr(xlErrRef  ): errText = "xlErrRef"
              case cvErr(xlErrValue): errText = "xlErrValue"
              case else             : errText = "?"
       end select

       debug.print("Cell at " & r & "," & c & " has the error value " & errText)
    else
       debug.print("Value of cell at " & r & "," & c & " is: " & cellVal)
    end if

end sub ' }
'
' Cell at 1,2 has the error value xlErrDiv0
' Cell at 2,2 has the error value xlErrNA
' Cell at 3,2 has the error value xlErrName
' Cell at 4,2 has the error value xlErrNull
' Cell at 5,2 has the error value xlErrNum
' Cell at 6,2 has the error value xlErrRef
' Cell at 7,2 has the error value xlErrValue
' Value of cell at 1,3 is: 42
