option explicit

sub main() ' {

   dim db as dao.database
   set db = application.currentDB

 '
 ' Check if we have transactions enabled:
 '
   if not db.properties("transactions") then ' {

      debug.print("Current database is not configured to use transactions")
      exit sub

   end if ' }

 '
 ' Drop table if it already exists:
 '
   if not isNull(dLookup("Name", "MSysObjects", "Name='tq84_trx'")) then db.execute("drop table tq84_trx")

 '
 ' Create the test table
 '
   db.execute("create table tq84_trx(id number, txt varchar(30))")

 '
 ' Insert a record out
 '
   db.execute("insert into tq84_trx values (1, 'Not within transaction')")

 '
 ' Begin first transaction
 '

   dao.DBEngine.beginTrans

 '
 ' Insert another record …
 '
   db.execute("insert into tq84_trx values (2, 'First transaction')")

 '
 ' … and commit it.
 '
   dao.DBEngine.commitTrans

 '
 ' Another transaction …
 '
   dao.DBEngine.beginTrans
   db.execute("insert into tq84_trx values (3, 'Second transaction')")

 '
 ' … but roll it back this time:
 '
   dao.DBEngine.rollback

 '
 ' Select values in table (to demonstrate that third
 ' record was rolled back).
 '
   dim rs as dao.recordSet
   set rs = db.openRecordset("tq84_trx")
   do while not rs.eof ' {
      debug.print (rs!id & " " & rs!txt)
      rs.moveNext
   loop

end sub ' }
