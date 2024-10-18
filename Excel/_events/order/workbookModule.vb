option explicit

private sub workbook_open()
    evt "workbook_open"            , me.name
end sub

private sub workbook_activate()
    evt "workbook_activate"        , me.name
end sub

private sub workbook_windowActivate(byVal wn as window)
    evt "workbook_windowActivate"  ,  wn.activeSheet.name
end sub

private sub workbook_windowDeActivate(byVal wn as window)
    evt "workbook_windowDeactivate", wn.activeSheet.name
end sub

private sub workbook_WindowResize(byVal wn as window)
    evt "workbook_windowResize"    ,  wn.activeSheet.name & " - " & wn.width & "x" & wn.height & " @ " & wn.top & "," &wn.left
end sub

private sub workbook_sheetActivate(byVal sh as object)
    evt "workbook_sheetActivate" , sh.name
end sub

private sub workbook_sheetDeactivate(byVal sh as object)
    evt "workbook_sheetDeactivate" , sh.name
end sub

private sub workbook_deactivate()
    evt "workbook_deactivate"      , me.name
end sub

private sub workbook_beforeClose(cancel as boolean)
    evt "workbook_beforeClose"     , me.name
end sub

private sub workbook_sheetChange(ByVal Sh As object, byVal target as range)
    evt "workbook_sheetChange"     , sh.name & ", target.address = " & target.address
end sub
