option explicit

sub main() ' {

    dim r as long
    r = 1

    cells(r,  2) = "-1.0"
    cells(r,  3) = "-0.8"
    cells(r,  4) = "-0.6"
    cells(r,  5) = "-0.4"
    cells(r,  6) = "-0.2"
    cells(r,  7) = " 0.0"
    cells(r,  8) = " 0.2"
    cells(r,  9) = " 0.4"
    cells(r, 10) = " 0.6"
    cells(r, 11) = " 0.8"
    cells(r, 12) = " 1.0"

    range(cells(r,2), cells(r,12)).horizontalAlignment  = xlCenter
    range(columns(2), columns(12)).columnWidth = 5

    drawThemeColor xlThemeColorAccent1           , "xlThemeColorAccent1"           , r
    drawThemeColor xlThemeColorAccent2           , "xlThemeColorAccent2"           , r
    drawThemeColor xlThemeColorAccent3           , "xlThemeColorAccent3"           , r
    drawThemeColor xlThemeColorAccent4           , "xlThemeColorAccent4"           , r
    drawThemeColor xlThemeColorAccent5           , "xlThemeColorAccent5"           , r
    drawThemeColor xlThemeColorAccent6           , "xlThemeColorAccent6"           , r
    drawThemeColor xlThemeColorDark1             , "xlThemeColorDark1"             , r
    drawThemeColor xlThemeColorDark2             , "xlThemeColorDark2"             , r
    drawThemeColor xlThemeColorFollowedHyperlink , "xlThemeColorFollowedHyperlink" , r
    drawThemeColor xlThemeColorHyperlink         , "xlThemeColorHyperlink"         , r
    drawThemeColor xlThemeColorLight1            , "xlThemeColorLight1"            , r
    drawThemeColor xlThemeColorLight2            , "xlThemeColorLight2"            , r

    columns(1).autoFit

    activeWorkbook.saved = true

end sub ' }


sub drawThemeColor(col as xlThemeColor, nam as string, byRef r as long) ' {

    r = r + 1

    cells(r, 1) = nam

    range(cells(r,2), cells(r,12)).interior.themeColor = col

    cells(r, 2).interior.tintAndShade =  -1.0
    cells(r, 3).interior.tintAndShade =  -0.8
    cells(r, 4).interior.tintAndShade =  -0.6
    cells(r, 5).interior.tintAndShade =  -0.4
    cells(r, 6).interior.tintAndShade =  -0.2
    cells(r, 7).interior.tintAndShade =   0.0
    cells(r, 8).interior.tintAndShade =   0.2
    cells(r, 9).interior.tintAndShade =   0.4
    cells(r,10).interior.tintAndShade =   0.6
    cells(r,11).interior.tintAndShade =   0.8
    cells(r,12).interior.tintAndShade =   1.0


end sub ' }
