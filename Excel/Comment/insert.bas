option explicit

sub main() ' {

    dim cl as range
    dim cm as comment
    dim sh as shape

    set cl = cells(2,2)
    cl.value = "foo"

    set cm = cl.addComment
  '
  ' Note: the text of the comment is assigned with a method, not
  '       a property:
  '
    cm.text("A comment that" & chr(10) & "should describe foo.")
    cm.visible = true

    set sh = cm.shape

    sh.height = 25
    sh.width  = 90

    activeWorkBook.saved = true

end sub ' }
