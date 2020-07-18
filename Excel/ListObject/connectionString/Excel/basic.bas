option explicit

sub main() ' {

    dim curPath as string
    curPath = thisWorkbook.path & chr$(92)  ''' chr$(92) is the backslash.

    dim pathToSourceWorkbook As String
    pathToSourceWorkbook = curPath & "workbook-with-src-data.xlsx"

    createSourceWorksheet pathToSourceWorkbook

    insertListObjectWith  _
       source       :=  "OLEDB;provider=Microsoft.ACE.OLEDB.16.0;data source=" & pathToSourceWorkbook & ";extended properties=""excel 12.0;hdr=yes""", _
       sqlStatement :=  "select [Col two], [Col three] from [srcTable] where [Col one] = 'Baz'" , _
       destCell     := cells(2,2)

end sub ' }

sub insertListObjectWith( source as string, sqlStatement as string, destCell as range) ' {

    dim listObj as listObject

    set listObj = activeSheet.listObjects.add( _
        sourceType  := xlSrcExternal         , _
        source      := array(source)         , _
        destination := destCell)

    with listObj ' {

        .displayName = "Data_from_other_worksheet" ' Must not contain white spaces

         with .queryTable ' {

'            .adjustColumnWidth      = true                  ' True is default anyway

             .commandType            = xlCmdSql
             .commandText            = array(sqlStatement)
'            .rowNumbers             = false

             .refreshOnFileOpen      = false                 ' Get newest data when worksheet is opened (Default is false)
             .backgroundQuery        = true                  ' Update data asynchronously
             .refreshStyle           = xlInsertDeleteCells   ' Partial rows are inserted or deleted to match the exact number of rows required for the new recordset.
             .saveData               = true
             .refreshPeriod          = 0                     ' Refresh period in minuts. 0 disables refreshing.
             .preserveColumnInfo     = true                  ' Preserve sorting, filtering, and layout information when data is refreshed.


             .refresh backgroundQuery := false               ' Refresh the data NOW.

         end with ' }

      '
      '  Apparently, date format of source table is not automatically transferred
      '  to destination table. So, we have to explicitely define it:
      '
        .listColumns("Col Three").range.numberFormat = "m/d/yyyy"

    end with ' }

end sub ' {

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
