option explicit

sub main() ' {

    dim connectionString as string

    connectionString =                            _
       "OLEDB;"                                 & _
       "provider=Microsoft.ACE.OLEDB.12.0;"     & _
       "data source=" & thisWorkbook.path & ";" & _
       "extended Properties=text"

    dim destTable as listObject

    set destTable = activeSheet.listObjects.add( _
       sourceType  := xlSrcExternal            , _
       source      := connectionString         , _
       destination := cells(2, 2))

    with destTable.queryTable

        .commandType     = xlCmdSql
        .commandText     = array("select * from [tab.csv]")
        .backgroundQuery = true

        .refresh backgroundQuery := false

    end With

end sub ' }
