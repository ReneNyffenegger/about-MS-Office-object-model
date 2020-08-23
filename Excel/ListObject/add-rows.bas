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

    table.listRows.add.range.value = array("foo", 42, #2020-03-05#)
    table.listRows.add.range.value = array("bar", 99, #2017-10-17#)
    table.listRows.add.range.value = array("baz",  5, #2018-03-21#)

end sub ' }
