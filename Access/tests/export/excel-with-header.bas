' In order to prevent compile error »User-defined type not defined« add the reference to the Excel Object Library in the immediate window:
'
'          call application.VBE.activeVBProject.references.addFromGuid("{00020813-0000-0000-C000-000000000046}", 1, 8)
'

option explicit

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

    dropTableIfExists db, "tq84_data"
    db.execute("create table tq84_data(foo number, bar varchar(20), baz date)")

    db.execute("insert into tq84_data(foo, bar, baz) values (1   , ""one""  , now()                )")
    db.execute("insert into tq84_data(foo, bar, baz) values (2   , ""two""  , null                 )")
    db.execute("insert into tq84_data(foo, bar, baz) values (null, ""three"", #2019-01-05 12:34:56#)")
    db.execute("insert into tq84_data(foo, bar, baz) values (4   ,   null   , #2022-04-18 16:50:27#)")

    dim excelFileName as string
    excelFileName = environ$("TEMP") & "\access-to-excel-export.xlsx"
    doCmd.transferSpreadsheet acExport, acSpreadsheetTypeExcel12XML, "tq84_data", excelFileName, true

    excelMakeFirstRowHeader excelFileName

end sub ' }

sub excelMakeFirstRowHeader(excelFileName as string) ' {

  dim xls as new excel.application
  dim wkb as     excel.workbook

  set wkb = xls.workbooks.open(excelFileName)
  wkb.activeSheet.rows("2:2").select
  xls.activeWindow.freezePanes = true

  wkb.save
  wkb.close

end sub ' }

sub dropTableIfExists(db as dao.database, tableName as string) ' {
  on error goto err_
    db.execute("drop table " & tableName)
    exit sub
  err_:
    if err.number = 3376 then
     '
     ' Ignore »Table … does not exist«.
     '
       exit sub
    end if

    err.raise err.number, err.source, err.description

end sub ' }
