option explicit

sub main() ' {
 '
 '  Text is converted to number: leading zeroes are removed
 '  and number is right aligned:
 '
    cells(1,1) = "0001"

 '
 '  Text is inserted as number THEN converted to string
 ' (which still removes leading zeroes)
 '
    cells(2,1) = "0002"
    cells(2,1).numberFormat = "@"

 '
 '  Format of cell is changed to text THEN text is inserted.
 '  This keeps leading zeroes but also has green triangle that
 '  indicates an error in the cell
 '
    cells(3,1).numberFormat = "@"
    cells(3,1) = "0003"

 '
 '  Remove green triangle by setting errors(…).ignore
 '  property to true:
 '
    cells(4,1).numberFormat = "@"
    cells(4,1) = "0004"
    cells(4,1).errors(xlNumberAsText).ignore = true

    activeWorkBook.saved = true

end sub ' }
