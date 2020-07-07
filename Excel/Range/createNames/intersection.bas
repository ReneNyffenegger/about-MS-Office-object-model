option explicit

sub main() ' {

    testData

    cells(2,2).currentRegion.createNames top := true, left := true

    dim nameCol, nameRow as string
    nameCol = "valThree"
    nameRow = "fourthRow"

    call highlightCol(nameCol)
    call highlightRow(nameRow)

    cells(9, 3).formula = "= ""Value of " & nameCol & " in " & nameRow & " is: "" & cell(""contents"", " & nameCol & " " & nameRow & ")"

end sub ' }


sub highlightCol(byVal name as string) ' {

    dim rng as range
    set rng = range (name)

    rng.font.color  = rgb(230, 110, 0)
    rng.font.bold = true

end sub ' }


sub highlightRow(byVal name as string) ' {

    dim rng as range
    set rng = range (name)

    rng.interior.color = rgb(255, 220, 85)

end sub ' }


sub testData() ' {

                                 cells(2, 3) = "valOne"  : cells(2, 4) = "valTwo" : cells(2, 5) = "valThree" : cells(2, 6) = "valFour"
     cells(3, 2) = "firstRow"  : cells(3, 3) =      18   : cells(3, 4) =      22  : cells(3, 5) =         9  : cells(3, 6) =       45
     cells(4, 2) = "secondRow" : cells(4, 3) =       4   : cells(4, 4) =      57  : cells(4, 5) =        78  : cells(4, 6) =       16
     cells(5, 2) = "thirdRow"  : cells(5, 3) =      33   : cells(5, 4) =      43  : cells(5, 5) =        25  : cells(5, 6) =       69
     cells(6, 2) = "fourthRow" : cells(6, 3) =      25   : cells(6, 4) =      32  : cells(6, 5) =        74  : cells(6, 6) =       52
     cells(7, 2) = "fifthRow"  : cells(7, 3) =       9   : cells(7, 4) =      49  : cells(7, 5) =        61  : cells(7, 6) =       62

     range(cells(2,3), cells(2, 6)).font.bold = true
     range(cells(3,2), cells(7, 2)).font.bold = true

     range(columns(2), columns(6)).autofit

end sub ' }
