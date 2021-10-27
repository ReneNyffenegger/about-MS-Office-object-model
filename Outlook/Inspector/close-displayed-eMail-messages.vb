option explicit

sub closeDisplayedEmailMessages() ' {

    dim ins as inspector
    for each ins in application.inspectors

        if ins.currentItem.class = olMail then
           ins.currentItem.close olDiscard
        end if

    next ins

end sub ' }
