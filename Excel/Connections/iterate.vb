option explicit

sub t

    dim conn     as WorkbookConnection
    dim connType as string

    for each conn in activeWorkbook.connections

        connType = switch (                                            _
          conn.type = xlConnectionTypeDATAFEED , "Data Feed",          _
          conn.type = xlConnectionTypeMODEL    , "PowerPivot Model",   _
          conn.type = xlConnectionTypeNOSOURCE , "No source",          _
          conn.type = xlConnectionTypeODBC     , "ODBC",               _
          conn.type = xlConnectionTypeOLEDB    , "OLEDB",              _
          conn.type = xlConnectionTypeTEXT     , "Text",               _
          conn.type = xlConnectionTypeWEB      , "Web",                _
          conn.type = xlConnectionTypeWORKSHEET, "Worksheet",          _
          conn.type = xlConnectionTypeXMLMAP   , "XML MAP"             _
        )

        debug.print conn.name & " (" & connType & ")"

        if     conn.type = xlConnectionTypeOLEDB then
               debug.print "  Command type: " & conn.oledbConnection.commandType

               if conn.oledbConnection.commandType = xlCmdSql then
                  debug.print "  Command text: " & conn.oledbConnection.commandText
               end if
               debug.print "  Connection:   " & conn.oledbConnection.connection

               debug.print ""

        elseif conn.type = xlConnectionTypeXMLMAP then

               dim rng as range
               set rng = conn.ranges(1)

               debug.print "  " & rng.parent.name & "!" & rng.address
               debug.print ""

        end if


    next conn

end sub
