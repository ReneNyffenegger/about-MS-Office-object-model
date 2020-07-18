option explicit

sub main() ' {

    dim curPath as string
    curPath = thisWorkbook.path & chr$(92)  ''' chr$(92) is the backslash.

    dim pathToSourceWorkbook As String
    pathToSourceWorkbook = curPath & "workbook-with-src-data.xlsx"

    createSourceWorksheet pathToSourceWorkbook

    dim connectionString as string
    connectionString = "oledb;provider=Microsoft.ACE.OLEDB.16.0;" & _
             "data source=" & pathToSourceWorkbook          & ";" & _
             "extended properties=""excel 12.0;hdr=yes"""

    dim wbconn as workbookConnection
    Set wbconn = activeWorkbook.connections.add2                                  ( _
       name             := "connection to other excel sheet"                      , _
       Description      := "this connection was just created for testing purposes", _
       connectionString :=  connectionString                                      , _
       commandText      := "select * from [srcTable]"                             , _
       lCmdType         :=  xlCmdSql)


end sub ' }

sub createSourceWorksheet(fileName as string) ' {
  '
  '  Delete source workbook file if it alread exists.
  '
    if dir(fileName) <> "" then ' {
       kill fileName
    end if ' }

    dim otherWorkbook as workbook
    set otherWorkbook = workbooks.add

    with otherWorkbook ' {

      dim firstCell as range

      with .sheets(1) ' {

        dim r as long : r = 3
        set firstCell = .cells(r,2)

       .range( .cells(r, 2), .cells(r, 4) ).value = array("Col one", "Col two", "Col three"  ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Baz"    ,       42 , #2020-03-03# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Bar"    ,       99 , #2018-05-17# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Baz"    ,   123456 , #2019-11-13# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Foo"    ,      518 , #2018-07-19# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Baz"    ,      219 , #2014-10-02# ) : r = r + 1
       .range( .cells(r, 2), .cells(r, 4) ).value = array("Foo"    ,       21 , #2015-09-09# )

    '
    '   Name a source data range
    '
       .range( firstCell, .cells(r,4) ).name = "srcTable"

       .usedRange.columns.autoFit

      end with ' }

     .saveAs                            _
        fileName   := fileName,         _
        fileFormat := xlOpenXMLWorkbook

     .close

    end with ' }

end sub ' }
