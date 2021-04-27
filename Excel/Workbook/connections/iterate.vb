option explicit

sub iterateOverWorkbookConnections() ' {

    dim con as workbookConnection

    for each con in activeWorkbook.connections

        dim tp_ as long
        dim tp  as string

        tp_ = con.type

        if      tp_ = xlConnectionTypeOLEDB   then tp = "OLEDB"    _
        else if tp_ = xlConnectionTypeXMLMAP  then tp = "XML Map"  _
        else                                       tp = "?"

        debug.print con.name & ", type = " & tp & ", description = " & con.description

        dim rng as range
        for each rng in con.ranges
            debug.print "  " & rng.parent.name & ": " & rng.address
        next rng

        debug.print

        if     tp_ = xlConnectionTypeOLEDB then

           with con.oledbConnection
           debug.print "  connection  = " & .connection
           debug.print "  commandText = " & .commandText
           end with

        elseif tp_ = xlConnectionTypeXMLMAP then

        ' Nothing to see here?

        end if
        debug.print

    next con

end sub ' }
