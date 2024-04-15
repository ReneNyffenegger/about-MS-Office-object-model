option explicit

sub main()

   dim formula_m  as string
   dim query_name as string

   formula_m = "let result =              " & _
               "  Table.FromColumns( {    " & _
               "  {          42  ,            99   ,      7   }," & _
               "  { ""forty-two"", ""ninety-nine"", ""seven"" } " & _
               "}, {                      " & _
               "   ""num"", ""txt"" })    " & _
               "in                        " & _
               "  result"

 ' query_name = "qry"

   dim query as workbookQuery
   set query = activeWorkbook.queries.add(     _
           name     :=  "qry"                , _
           formula  :=  formula_m)

   dim connectionString as string
   connectionString = "OLEDB;"                             & _
                      "Provider=Microsoft.Mashup.OleDb.1;" & _
                      "Data Source=$Workbook$;"            & _
                      "Location=" & query.name & ";"       & _
                      "Extended Properties="""""

   dim destTable as listObject
   set destTable = activeSheet.listObjects.add( _
       sourceType  := xlSrcExternal            , _
       source      := connectionString         , _
       destination := activeSheet.cells(2,2))

   with destTable.queryTable
        .commandType            = xlCmdSql
        .commandText            = array("select * from [" & query.name & "]")
        .refresh backgroundQuery := false
    end with

end sub
