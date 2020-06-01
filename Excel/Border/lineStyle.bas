dim r as long
public sub main() ' {

    r = 2

    drawBorders "xlContinuous"   , xlContinuous
    drawBorders "xlDash"         , xlDash
    drawBorders "xlDashDot"      , xlDashDot
    drawBorders "xlDashDotDot"   , xlDashDotDot
    drawBorders "xlDot"          , xlDot
    drawBorders "xlDouble"       , xlDouble
    drawBorders "xlLineStyleNone", xlLineStyleNone
    drawBorders "xlSlantDashDot" , xlSlantDashDot

    columns(1).columnWidth = 3
    columns(2).autofit

    activeWorkbook.saved = true

end sub ' }

private sub drawBorders(name as string, style as xlLineStyle) ' {

    dim cell as range
    dim bord as border

    set cell = cells(r, 2)
    cell.borders.lineStyle = style

    cell.value = name
    r = r+2

end sub ' }
