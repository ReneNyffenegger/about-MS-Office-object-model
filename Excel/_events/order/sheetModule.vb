option explicit

private sub worksheet_activate()
    debug.print me.name & ": worksheet_activate"
end sub

private sub worksheet_deactivate()
    debug.print me.name & ": worksheet_deactivate"
end sub

private sub worksheet_change(byVal target as range)
    debug.print me.name & ": worksheet_change, target = " & target.range
end sub

private sub worksheet_selectionChange(byVal target as range)
    debug.print me.name & ": worksheet_selectionChange, target = " & target.range
end sub
