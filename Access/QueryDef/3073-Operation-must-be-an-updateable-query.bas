option explicit

sub test_3073() ' {

   dim db as dao.database
   set db = application.currentDB

 '
 ' Drop table if it already exists:
 '
   if not isNull(dLookup("Name", "MSysObjects", "Name='tab_3073'")) then db.execute("drop table tab_3073")

 '
 ' Create table …
 '
   db.execute("create table tab_3073(id number primary key, nr long, tx varchar(20))")

 '
 ' … and insert some records
 '
   db.execute("insert into tab_3073 values (1, 22, 'twenty-two')")
   db.execute("insert into tab_3073 values (2, 40, 'fourty'    )")
   db.execute("insert into tab_3073 values (3, 11, 'eighteen'  )")
   db.execute("insert into tab_3073 values (4, 21, 'twenty'    )")


   dim stmt as dao.queryDef
   set stmt = db.createQueryDef("", "parameters nr long, tx varchar(20); update tab_3073 set nr = [nr] where tx = [tx]")

 '
 ' ??? The following line(s) cause Run-time error 3073 »Operation must use an updateable query.« ???
 '
 '     Why, oh why am I cursed to use access in this project?
 '
   stmt.parameters!nr = 18 : stmt.parameters!tx = "eighteen" : stmt.execute
   stmt.parameters!nr = 20 : stmt.parameters!tx = "twenty"   : stmt.execute

 '
 ' Select values in table (code doesn't reach here…)
 '
   dim rs as dao.recordSet
   set rs = db.openRecordset("tab_3073")
   do while not rs.eof ' {
      debug.print (rs!nr & " " & rs!tx)
      rs.moveNext
   loop ' }

end sub ' }
