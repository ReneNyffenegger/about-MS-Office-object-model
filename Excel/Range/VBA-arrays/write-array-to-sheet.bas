option explicit

sub main() ' {

    dim rngRowNames as range
    dim rngColNames as range

    set rngRowNames = writeArrayToSheet(3, 2, array("1st row", "2nd row", "3rd row", "4th row"), false)
    set rngColNames = writeArrayToSheet(2, 3, array("val one", "val two", "val three"         ), true )

    rngRowNames.interior.color = rgb(240, 200, 170)
    rngColNames.font.bold      = true

end sub ' }

function writeArrayToSheet(r as long, c as long, ary as variant, optional horizontal as boolean = false, optional ws as worksheet = nothing) as range ' {

    if ws is nothing then
       set ws = activeSheet
    end if

    dim szArray_minusOne as long
    szArray_minusOne = uBound(ary) - lBound(ary)

    with ws ' {

        if horizontal then
           set writeArrayToSheet = .range(.cells(r, c), .cells(r, c+szArray_minusOne))
           writeArrayToSheet.value = ary
        else
           set writeArrayToSheet = .range(.cells(r, c), .cells(r+szArray_minusOne, c))
           writeArrayToSheet.value = application.worksheetFunction.transpose(ary)
        end if

    end with ' }

end function ' }
