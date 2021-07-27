option explicit

dim curDataRow  as long
public rngTestData as range

sub main() ' {

    fillTestData

  '
  '                 Sort on second column: -------+
  '                                               |
  '                                               V
    rngTestData.sort key1 := rngTestData.offset(0,1).resize(1,1), order1 := xlAscending, header := xlNo

end sub ' }

sub fillTestData() ' {

    curDataRow = 3

    addDataRow  1, "one"
    addDataRow  2, "two"
    addDataRow  3, "three"
    addDataRow  4, "four"
    addDataRow  5, "five"
    addDataRow  6, "six"
    addDataRow  7, "seven"
    addDataRow  8, "eight"
    addDataRow  9, "nine"
    addDataRow 10, "ten"

    set rngTestData = range(cells(3,3), cells(curDataRow-1, 4))

end sub ' }

sub addDataRow(numVal as long, numText as string) ' {

    cells(curDataRow, 3) = numVal
    cells(curDataRow, 4) = numText

    curDataRow = curDataRow + 1

end sub ' }
