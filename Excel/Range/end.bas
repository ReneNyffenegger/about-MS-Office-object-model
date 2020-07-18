option explicit

sub main() ' {

    activeSheet.usedRange.clearFormats
    activeSheet.usedRange.clearContents

    cells( 7, 1) = "A"
    cells( 9, 1) = "B"
    cells( 9, 3) = "C"
    cells( 9, 7) = "D"
    cells( 6, 7) = "E"
    cells( 3, 7) = "F"
    cells( 3, 5) = "G"
    cells( 7, 5) = "H"

    cells( 7, 2).select

    move xlToLeft
    move xlDown
    move xlToRight
    move xlToRight
    move xlUp
    move xlUp
    move xlToLeft
    move xlDown

    activeSheet.usedRange.columns.autoFit

end sub ' }

sub move(direction as xlDirection) ' {

    dim currentCell as range : set currentCell = selection
    dim nextCell    as range : set nextCell    = currentCell.end(direction)

    range(currentCell, nextCell).interior.color = rgb(255, 200,  40)

    nextCell.select

end sub ' }
