option explicit

sub main()

   dim rmd as reminder
   for each rmd in application.reminders
        debug.print rmd.caption & ": " & rmd.isVisible & ", " & typename(rmd.item)
   next rmd

end sub
