option explicit

sub main() ' {

    dim actCell     as range
    dim actWindow   as window
    dim actSheet    as worksheet
    dim actWorkbook as workbook

    set actWorkbook = application.activeWorkbook
    set actSheet    = application.activeSheet
    set actCell     = application.activeCell
    set actWindow   = application.activeWindow

    debug.print("Active workbook: " & actWorkbook.name )
    debug.print("Active sheet:    " & actSheet.name    )
    debug.print("Active cell:     " & actCell.address  )
    debug.print("Active window:   " & actWindow.caption)

end sub ' }
