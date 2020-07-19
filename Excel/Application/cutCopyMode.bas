option explicit

sub main() ' {


    application.width  = 500
    application.height = 400

  [ b2:e4 ] = [{ "abc","def",2490,6104 ; "ghi","jkl",2828,96344 ; "mno","pqr",19041,940902 }]

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "Going to select a region"

  [ c3:d4 ].select

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "Going to simulate ctrl-c (copy)"

    selection.copy

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "Going to select the destination cell"

  [ b5    ].select

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "Going to simulate ctrl-v (paste)"

    activeSheet.paste

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "Going to 'unselect' selected from region"

    application.cutCopyMode = false

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "Going to select another region"

  [ e2:e3 ].select

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "Going to simulate ctrl-x (cut)"

    selection.cut

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "Going to select the destination cell"

  [ a3    ].select

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "Going to simulate ctrl-v (paste)"

    activeSheet.paste

    msgBox "cutCopyMode = " & cutCopyModeToString & chr(10) & "finished"

end sub ' }

function cutCopyModeToString() as string ' {

    select case application.cutCopyMode
       case false : cutCopyModeToString = "false"
       case xlCopy: cutCopyModeToString = "xlCopy"
       case xlCut : cutCopyModeToString = "xlCut"
       case else  : cutCopyModeToString = "???"
    end select

end function ' }
