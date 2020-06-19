option explicit

public sub main() ' {

  '
  ' Create range object for 3rd row, 2nd column
  ' to 9th row, 5th column:
  '
    dim rng as range
    set rng = range(    _
       cells(3, 2),     _
       cells(9, 5)      _
    )

  '
  ' Use range object to modify some properties
  ' of all cells that belong to the range:
  '
    rng.formula      = "=rand()"
    rng.numberFormat = "0.000"

    with rng.font
        .name = "Lucida Console"
        .size =  10
    end with

    rng.columns.autoFit

    activeWorkbook.saved = true

end sub ' }
