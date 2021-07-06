option explicit

sub main() ' {
    createCellsWithErrors

    dim r as long

    for r = 1 to 9 ' {
        checkIfCellContainsError  r
    next r ' }

    activeSheet.usedRange.columns.autoFit

    cells(11, 11).select

end sub ' }

sub createCellsWithErrors() ' {

'   xlErrDiv0
'   -----------------------------------------
    cells(1, 3) =  42
    cells(1, 4) =   0
    cells(1, 2).formulaR1C1 = "= RC[1] / RC[2]"

'   xlErrNA
'   -----------------------------------------
    cells(2, 3) = "foo"
    cells(2, 4) = "bar"
    cells(2, 5) = "baz"
    cells(2, 6) = "n.a."
    cells(2, 2).formulaR1C1 = "=match(RC[4],RC[1]:RC[3], 0)"

'   xlErrName
'   -----------------------------------------
    cells(3, 2).formulaR1C1 = "= inexistingFunction()"

'   xlErrNull
'   -----------------------------------------
    cells(4, 2).formulaR1C1 = "= sum(RC[1]:RC[2] RC[3]:RC[4])" ' Use intersect operator (space) on two ranges that have no cell in common

'   cells(5, 1) = "xlErrNum"
    cells(5, 3).formulaR1C1 = "= 17.9769313486231 * 1e307"     ' Close to the maximum double precision floating-point value
    cells(5, 2).formulaR1C1 = "= 17.9769313486232 * 1e307"     ' Over the maximum double precision floating-point value

'   xlErrRef
'   -----------------------------------------
    cells(6, 9) = "some value that is going to be deleted"
    cells(6, 2).formulaR1C1 = "= RC9"
    columns(9).delete

'   xlErrValue
'   -----------------------------------------
    cells(7, 2).formulaR1C1 = "= ""foo"" + ""bar"""            ' cannot add two strings…

'   xlEmptyCellReferences
'   -----------------------------------------
    cells(8, 3).value       = "17"
    cells(8, 2).formulaR1C1 = "= RC[1] + RC[2]"                ' Some referenced cells in the formula are empty

'   xlNumberAsText
'   -----------------------------------------
    cells(9, 2).value       = "'99"                            ' Prepend number with apostrophe


end sub ' }

sub checkIfCellContainsError(r as long) ' {

    dim cell as range : set cell = cells(r, 2)

    dim cellVal as variant
    cellVal = cell.value

    dim errText as string


    if cell.errors(xlEmptyCellReferences    ).value then errText = "xlEmptyCellReferences"
    if cell.errors(xlInconsistentFormula    ).value then errText = "xlInconsistentFormula"
    if cell.errors(xlInconsistentListFormula).value then errText = "xlInconsistentListFormula"
    if cell.errors(xlListDataValidation     ).value then errText = "xlListDataValidation"
    if cell.errors(xlMisleadingFormat       ).value then errText = "xlMisleadingFormat"
    if cell.errors(xlNumberAsText           ).value then errText = "xlNumberAsText"
    if cell.errors(xlOmittedCells           ).value then errText = "xlOmittedCells"
    if cell.errors(xlTextDate               ).value then errText = "xlTextDate"
    if cell.errors(xlUnlockedFormulaCells   ).value then errText = "xlUnlockedFormulaCells"

    if cell.errors(xlEvaluateToError        ).value then ' {
       errText = errText & ", " & "xlEvaluateToError"

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

      if varType(cellVal) <> vbError then ' {
         msgBox "Unepxected: varType(cellVal) <> vbError, but xlEvaluateToError"
      end if ' }

    end if ' }

    cells(r, 1) = errText

end sub ' }
