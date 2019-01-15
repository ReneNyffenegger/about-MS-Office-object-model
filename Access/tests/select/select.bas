' In order to prevent compile error »User-defined type not defined« add the reference to ADODB in the immediate window:
'
'          call application.VBE.activeVBProject.references.addFromGuid("{B691E011-1797-432E-907A-4D8C69339129}", 6, 1)
'
option explicit

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

    cleanUpLastRun   db
    createTables     db
    insertValues     db

    selectValues_DAO db
    selectValues_ADO currentProject.connection

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

sub cleanUpLastRun(db as dao.database) ' {

    call dropTableIfExists(db, "tq84_data"        )
    call dropTableIfExists(db, "tq84_lookUp_one"  )
    call dropTableIfExists(db, "tq84_lookUp_two"  )
    call dropTableIfExists(db, "tq84_lookUp_three")

end sub ' }

sub createLookupTable(db as dao.database, tableName as string) ' {

    db.execute(   _
      "create table " & tableName  & "("                      & _
      "  id     long primary key,"                            & _
      "  txt    varchar(10)      "                            & _
      ")")

end sub ' }

sub createTables(db as dao.database) ' {

    call createLookupTable(db, "tq84_lookUp_one"   )
    call createLookupTable(db, "tq84_lookUp_two"   )
    call createLookupTable(db, "tq84_lookUp_three" )

    db.execute(   _
      "create table tq84_data ("                          & _
      "  txt    varchar(20), "                            & _
      "  id_1   long  references tq84_lookUp_one,"        & _
      "  id_2   long  references tq84_lookUp_two,"        & _
      "  id_3   long  references tq84_lookUp_three"       & _
      ")")

end sub ' }

function insertStatementForLookupTable(db as dao.database, tableName as string) as dao.queryDef ' {

     set insertStatementForLookupTable  = db.createQueryDef("",   _
       "parameters "         & _
       "  id  number  , "    & _
       "  txt varchar(10); " & _
       "insert into " & tableName & "(id, txt) values ([id], [txt]) ")

end function ' }

sub insertLookupValues(insertStatement as dao.queryDef, id as long, txt as string) ' {

    insertStatement!id  = id
    insertStatement!txt = txt

    insertStatement.execute

end sub ' }

sub insertDataValues(insertStatement as dao.queryDef, txt as string, id_1 as long, id_2 as long, id_3 as long) ' {

    insertStatement!txt  = txt
    insertStatement!id_1 = id_1
    insertStatement!id_2 = id_2
    insertStatement!id_3 = id_3

    insertStatement.execute

end sub ' }

sub insertValues(db as dao.database) ' {

    dim insertStatement_lookUp_one   as dao.queryDef
    dim insertStatement_lookUp_two   as dao.queryDef
    dim insertStatement_lookUp_three as dao.queryDef

    set insertStatement_lookUp_one   = insertStatementForLookupTable(db, "tq84_lookUp_one"  )
    set insertStatement_lookUp_two   = insertStatementForLookupTable(db, "tq84_lookUp_two"  )
    set insertStatement_lookUp_three = insertStatementForLookupTable(db, "tq84_lookUp_three")

    call insertLookupValues(insertStatement_lookUp_one  , 1, "one"   )
    call insertLookupValues(insertStatement_lookUp_one  , 2, "two"   )
    call insertLookupValues(insertStatement_lookUp_one  , 3, "three" )

    call insertLookupValues(insertStatement_lookUp_two  , 4, "four"  )
    call insertLookupValues(insertStatement_lookUp_two  , 5, "five"  )
    call insertLookupValues(insertStatement_lookUp_two  , 6, "six"   )

    call insertLookupValues(insertStatement_lookUp_three, 7, "seven" )
    call insertLookupValues(insertStatement_lookUp_three, 8, "eight" )
    call insertLookupValues(insertStatement_lookUp_three, 9, "nine"  )


     dim insertStatement_data   as dao.queryDef
     set insertStatement_data = db.createQueryDef("",               _
       "parameters "                                   & _
       "  txt varchar(20),"                            & _
       "  id_1 number ,"                               & _
       "  id_2 number ,"                               & _
       "  id_3 number; "                               & _
       "insert into tq84_data(txt, id_1, id_2, id_3) " & _
       "values ([txt], [id_1], [id_2], [id_3])")

     call insertDataValues(insertStatement_data, "one five seven"  , 1, 5, 7)
     call insertDataValues(insertStatement_data, "three five eight", 3, 5, 8)
     call insertDataValues(insertStatement_data, "three four nine" , 3, 4, 9)
     call insertDataValues(insertStatement_data, "two six seven"   , 2, 6, 7)

end sub ' }

sub selectValues_DAO(db as dao.database) ' {

    dim stmt as queryDef
    set stmt = db.createQueryDef("",                              _
      "select "                                                 & _
      "  dt.txt as txt_data,        "                           & _
      "  l1.txt as txt_lookup_one,  "                           & _
      "  l2.txt as txt_lookup_two,  "                           & _
      "  l3.txt as txt_lookup_three "                           & _
      "from (("                                                 & _
      "  tq84_data         dt                      inner join " & _
      "  tq84_lookUp_one   l1 on dt.id_1 = l1.id ) inner join " & _
      "  tq84_lookUp_two   l2 on dt.id_2 = l2.id ) inner join " & _
      "  tq84_lookUp_three l3 on dt.id_3 = l3.id              ")

    dim rs as dao.recordSet
    set rs = stmt.openRecordSet

    debug.print("DAO:")
    do while not rs.eof ' {
       debug.print(rs!txt_data & ":  " & rs!txt_lookup_one & " - " & rs!txt_lookup_two & " - " & rs!txt_lookup_three)
       rs.moveNext
    loop ' }

end sub ' }

sub selectValues_ADO(cn as adodb.connection) ' {

    dim rs as new adodb.recordSet

    rs.open _
      "select "                                                 & _
      "  dt.txt as txt_data,        "                           & _
      "  l1.txt as txt_lookup_one,  "                           & _
      "  l2.txt as txt_lookup_two,  "                           & _
      "  l3.txt as txt_lookup_three "                           & _
      "from (("                                                 & _
      "  tq84_data         dt                      inner join " & _
      "  tq84_lookUp_one   l1 on dt.id_1 = l1.id ) inner join " & _
      "  tq84_lookUp_two   l2 on dt.id_2 = l2.id ) inner join " & _
      "  tq84_lookUp_three l3 on dt.id_3 = l3.id              " , _
      cn

    debug.print(""    )
    debug.print("ADO:")
    do until rs.eof ' {
       debug.print(rs!txt_data & ":  " & rs!txt_lookup_one & " - " & rs!txt_lookup_two & " - " & rs!txt_lookup_three)
       rs.moveNext
    loop ' }

end sub ' }
