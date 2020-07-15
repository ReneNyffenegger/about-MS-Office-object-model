option explicit

public sub main()

    activeSheet.usedRange.clearFormats
    activeSheet.usedRange.clearContents

    dim range_orig   as range
    dim range_offset as range

    set range_orig = range(cells(4, 3), cells(5, 5))

    with range_orig
        .value = "Orig"
        .font.color = rgb( 40, 220, 110)
    end with

  ' New range: 1 upward, 2 leftward
    set range_offset = range_orig.offset(-1, 2)

    range_offset.interior.color = rgb(255, 127, 30)

    range_orig.offset(columnOffset := -1).borderAround xlContinuous, xlThick, color := rgb(30, 110, 240)

    activeWorkbook.saved = true

end sub
