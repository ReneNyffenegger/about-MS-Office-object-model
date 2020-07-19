option explicit

sub main() ' {

  '
  ' Use a range to determine the coordinates where
  ' the button will be placed:
  '
    dim dest as range
    set dest = range( cells(4,2), cells(5,3) )

  '
  ' Create/add the button.
  '
    dim btn as button
    set btn  = activeSheet.buttons.add( left := dest.left, top := dest.top, width := dest.width, height := dest.height)
    btn.caption  = "Click me!"

  '
  ' Assign a macro (sub) to the button:
  '
    btn.onAction = "btnClicked"

end sub ' }

sub btnClicked() ' {

    msgBox "The button was apparently clicked"

end sub ' }
