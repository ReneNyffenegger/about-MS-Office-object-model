option explicit

sub main() ' {

    dim rngHeader as range
    set rngHeader = range(cells(2,2), cells(2,4))
    rngHeader.value = array("col one", "col two", "col three")

    dim table as listObject
    set table = activeSheet.listObjects.add(   _
       sourceType             := xlSrcRange  , _
       source                 := rngHeader   , _
       xlListObjectHasHeaders := xlYes)

    table.listRows.add.range.value = array("ABC", "forty-two"   , 42)
    table.listRows.add.range.value = array("PQR", "ninety-nine" , 99)
    table.listRows.add.range.value = array("PQR", "seven"       ,  7)
    table.listRows.add.range.value = array("IJK", "fourty"      , 40)
    table.listRows.add.range.value = array("ABC", "thirteen"    , 13)
    table.listRows.add.range.value = array("PQR", "seventy-two" , 72)
    table.listRows.add.range.value = array("IJK", "thirty-nine" , 39)
    table.listRows.add.range.value = array("XYZ", "sixty-eight" , 68)
    table.listRows.add.range.value = array("ABC", "twelve"      , 12)
    table.listRows.add.range.value = array("XYZ", "seventy-four", 72)
    table.listRows.add.range.value = array("IJK", "ninety-three", 93)
    table.listRows.add.range.value = array("PQR", "eighty-five" , 85)
    table.listRows.add.range.value = array("XYZ", "thirty-one"  , 31)
    table.listRows.add.range.value = array("IJK", "twenty"      , 20)

    activeSheet.usedRange.columns.autofit

    table.range.autoFilter                         _
          field     := 1                         , _
          criteria1 := array("ABC", "IJK", "XYZ"), _
          operator  := xlFilterValues

    table.range.autoFilter                         _
          field     := 3                         , _
          criteria1 :=">50"

end sub ' }
