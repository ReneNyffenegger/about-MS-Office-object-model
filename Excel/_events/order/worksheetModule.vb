option explicit

private sub worksheet_activate()
    evt "worksheet_activate", me.name
end sub

private sub worksheet_deactivate()
    evt "worksheet_deactivate", me.name
end sub

private sub worksheet_change(byVal target as range)
    evt "worksheet_change", me.name & ", target = " & target.address
end sub

private sub worksheet_selectionChange(byVal target as range)
    evt "worksheet_selectionChange", me.name & ", target = " & target.address
end sub
