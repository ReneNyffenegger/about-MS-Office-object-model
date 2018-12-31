option explicit

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

    createTable  db
    insertValues db
    selectValues db

 '
 '  Cleaning up
 '
    db.execute("drop table someTable")

end sub ' }

sub createTable(db as dao.database) ' {

    db.execute(   _
   "create table someTable (" & _
   "  colInt integer    ,  "  & _
   "  colDbl double     ,  "  & _
   "  colTxt varchar(10),  "  & _
   "  colDat date          "  & _
   ")")


end sub ' }

sub insertValues(db as dao.database) ' {

  '
  '  Create queryDef for insert statement:
  '

     dim stmt    as dao.queryDef
     set stmt = db.createQueryDef("",   _
       "parameters "                     & _
       "  parInt    integer    , "       & _
       "  parDbl    double     , "       & _
       "  parTxt    varchar(10), "       & _
       "  parDat    date       ; "       & _
       "  insert into someTable values ([parInt], [parDbl], [parTxt], [parDat])")

  '
  '  Prepare parameters:
  '

     dim valInt  as dao.parameter
     dim valDbl  as dao.parameter
     dim valTxt  as dao.parameter
     dim valDat  as dao.parameter

     set valInt = stmt.parameters("parInt")
     set valDbl = stmt.parameters("parDbl")
     set valTxt = stmt.parameters("parTxt")
     set valDat = stmt.parameters("parDat")

  '
  '  Insert first record
  '

     valInt = 42
     valDbl = 12.345
     valTxt ="foo"
     valDat = dateSerial(2010, 11, 12)

     stmt.execute

  '
  '  Insert seconds record
  '

     valInt =-28
     valDbl = 39.993
     valTxt ="bar"
     valDat = dateSerial(2000, 01, 02)

     stmt.execute

  '
  '  Insert third record
  '

     valInt = 99
     valDbl = null
     valTxt ="baz"
     valDat = null

     stmt.execute


end sub ' }

sub selectValues(db as dao.database) ' {

    dim stmt as queryDef
    set stmt = db.createQueryDef(""     , _
      "parameters valInt integer; " & _
      "select * from someTable where colInt > [valInt]")

    dim parInt as dao.parameter
    set parInt = stmt.parameters("valInt")
    parInt = 10


    dim rs as dao.recordSet
    set rs = stmt.openRecordSet

    do while not rs.eof

       debug.print(rs("colInt") & " | " & nz(rs("colDbl"), "NUL   ") & " | " & nz(rs("colTxt"), "NUL") & " | " & nz(rs("colDat"), "NUL      "))

       rs.moveNext
    loop

end sub ' }
