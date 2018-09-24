option explicit

sub main() ' {

    application.customizationContext = thisDocument

    application.keyBindings.add keyCode := buildKeyCode(wdKeyF8                            ), keyCategory := wdKeyCategoryCommand, command := "f8WasPressed"
    application.keyBindings.add keyCode := buildKeyCode(wdKeyF8, wdKeyControl              ), keyCategory := wdKeyCategoryCommand, command := "ctrlF8WasPressed"
    application.keyBindings.add keyCode := buildKeyCode(wdKeyF8, wdKeyShift                ), keyCategory := wdKeyCategoryCommand, command := "shiftF8WasPressed"
    application.keyBindings.add keyCode := buildKeyCode(wdKeyF8, wdKeyControl, wdKeyShift  ), keyCategory := wdKeyCategoryCommand, command := "ctrlShiftF8wasPressed"

end sub ' }

sub f8WasPressed() ' {
    debug.print "F8 was pressed"
end sub ' }

sub ctrlF8WasPressed() ' {
    debug.print "Ctrl+F8 was pressed"
end sub ' }

sub shiftF8WasPressed() ' {
    debug.print "Shift+F8 was pressed"
end sub ' }

sub ctrlShiftF8WasPressed() ' {
    debug.print "Ctrl+Shift+F8 was pressed"
end sub ' }
