option explicit

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

    cleanUpLastRun db
    createTables   db
    insertValues   db
    selectValues   db

end sub ' }

sub cleanUpLastRun(db as dao.database) ' {

    if not db.tableDefs("tq84_child" ) is nothing then db.execute("drop table tq84_child" )
    if not db.tableDefs("tq84_parent") is nothing then db.execute("drop table tq84_parent")

end sub ' }

sub createTables(db as dao.database) ' {

    db.execute(   _
      "create table tq84_parent ("                            & _
      "  id     long primary key,"                            & _
      "  txt    char(10)        "                             & _
      ")")

    db.execute(   _
      "create table tq84_child ("                             & _
      "  id_parent long       null references tq84_parent,"   & _
      "  txt    char(10)        "                             & _
      ")")

end sub ' }

sub insertValues(db as dao.database) ' {

     dim stmtParent   as dao.queryDef
     set stmtParent = db.createQueryDef("",   _
       "parameters "         & _
       "  id  number  , "    & _
       "  txt char(10); "    & _
       "insert into tq84_parent(id, txt) values ([id], [txt]) ")

     dim stmtChild   as dao.queryDef
     set stmtChild = db.createQueryDef("",   _
       "parameters "              & _
       "  id_parent  number, "    & _
       "  txt char(10)     ; "    & _
       "insert into tq84_child(id_parent, txt) values ([id_parent], [txt]) ")

     call insertValuesParent(stmtParent, 1, "one"    )
     call insertValuesParent(stmtParent, 2, "two"    )
     call insertValuesParent(stmtParent, 3, "three"  )

     call insertValuesChild (stmtChild , 1, "uno"    )
     call insertValuesChild (stmtChild , 1, "eins"   )
     call insertValuesChild (stmtChild , 3, "tre"    )
     call insertValuesChild (stmtChild , 4, "quattro") ' Note missing parent!

end sub ' }

sub insertValuesParent(stmt as dao.queryDef, id as long, txt as string) ' {

     stmt.parameters!id  = id
     stmt.parameters!txt = txt
     stmt.execute

end sub ' }

sub insertValuesChild(stmt as dao.queryDef, id_parent as long, txt as string) ' {

     stmt.parameters!id_parent = id_parent
     stmt.parameters!txt       = txt

   '
   ' Without dbFailOnError, the following stmt executes without
   ' throwing an error if id_parent does not refer to a record
   ' in tq84_parent - but the record is (obviously) not inserted!
   '
   ' Therefore, execute should, imho, always be used with
   ' dbFailOnError
   '
     stmt.execute  ' dbFailOnError

end sub ' }

sub selectValues(db as dao.database) ' {

    dim stmt as queryDef
    set stmt = db.createQueryDef("", _
      "select "                     & _
      "  p.txt as parent_txt, "     & _
      "  c.txt as child_txt   "     & _
      "from "                       & _
      "  tq84_parent p left join " & _
      "  tq84_child  c on p.id = c.id_parent")

    dim rs as dao.recordSet
    set rs = stmt.openRecordSet

    do while not rs.eof ' {
       debug.print(rs!parent_txt & ":  " & rs!child_txt)
       rs.moveNext
    loop ' }

end sub ' }
