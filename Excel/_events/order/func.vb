option explicit

sub evt(name as string, txt as string) ' {
    debug.print format(name & ":", "!" & string(26, "@")) & txt
end sub ' }


sub positionWindows()

    with application.windows("wbOne.xlsm")
        .left   =   12
        .width  =  323
        .top    =   42
        .height =  344
    end with

    with application.windows("wbTwo.xlsm")
        .left   =  339
        .width  =  323
        .top    =   42
        .height =  344
    end with

    with application.vbe.mainWindow
        .visible = true
        .left    =    6
        .width   =  883
        .top     =  516
        .height  =  500
    end with

end sub
