'
'    https://stackoverflow.com/a/53961227/180275
'
option explicit

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

  ' db.execute("drop table tab")

    createTable db
    insertValue db

end sub ' }

sub createTable(db as dao.database) ' {

    db.execute(   _
   "create table tab ( " & _
   "  id     guid,     " & _
   "  txt    char(60)  " & _
   ")")

   db.execute("alter table tab add constraint tab_pk primary key (id)")


end sub ' }

sub insertValue(db as dao.database) ' {

     dim stmt as dao.queryDef

     set stmt = db.createQueryDef("",   _
       "parameters "                  & _
       "  id     binary,   "          & _
       "  txt    char(60); "          & _
       "insert into tab values ([id], [txt]) ")


     dim parId   as dao.parameter
     dim parTxt  as dao.parameter

     set parId  = stmt.parameters("id" )
     set parTxt = stmt.parameters("txt")

     parId.value  =  GuidFromString("{936DA01F-9ABD-4D9D-80C7-02AF85C822A8}")
     parTxt.value = "Hello world."

     stmt.execute ' Access throws Runtime Error 3001 (Invalid argument)

end sub ' }
