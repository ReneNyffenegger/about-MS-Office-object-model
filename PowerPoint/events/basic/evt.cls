option explicit

public withEvents evtReceiver as application

private sub evtReceiver_slideShowBegin(byVal wn as slideShowWindow) ' {

    debug.print "slide show started at " & now()

end sub ' }
