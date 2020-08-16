option explicit

sub main() ' {

    dim srcRange as range
    set srcRange = testData

    srcRange.advancedFilter        _
      action      := xlFilterCopy, _
      copyToRange := cells(3,5)  , _
      unique      := true

    activeSheet.usedRange.columns.autoFit

end sub ' }

function testData() as range' {

    dim r as long : r = 3

    range(cells(r,2), cells(r,3)).value = array("Foo", 1) : r = r+1
    range(cells(r,2), cells(r,3)).value = array("Foo", 2) : r = r+1
    range(cells(r,2), cells(r,3)).value = array("Bar", 3) : r = r+1
    range(cells(r,2), cells(r,3)).value = array("Foo", 2) : r = r+1
    range(cells(r,2), cells(r,3)).value = array("Baz", 3) : r = r+1
    range(cells(r,2), cells(r,3)).value = array("Baz", 2) : r = r+1
    range(cells(r,2), cells(r,3)).value = array("Bar", 2) : r = r+1
    range(cells(r,2), cells(r,3)).value = array("Foo", 2) : r = r+1
    range(cells(r,2), cells(r,3)).value = array("Bar", 3)

    set testData = range(cells(3,2), cells(r,3))

end function ' }
