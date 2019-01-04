option explicit


type GUID ' {
  '
  '  Declared in rpcdce.h / included by rpc.h
  '
     Data1          as long
     Data2          as integer
     Data3          as integer
     Data4 (0 to 7) as byte
end  type ' }

declare function CoCreateGuid    lib "ole32" (pguid as GUID) as long
declare function StringFromGUID2 lib "ole32" (rguid as GUID, byVal lpOleChar as any, byVal cbmax as long) as long

function CoCreateGuid_ as GUID ' {

    if CoCreateGuid(CoCreateGuid_) <> 0 then
       MsgBox "Something went wrong with CoCreateGuid"
    end if

end function ' }

function StringFromGUID2_(rguid as GUID) as string ' {

    StringFromGUID2_ = space$(38)

    call StringFromGUID2 (rguid, strPtr(StringFromGUID2_), 38*2)

end function ' }

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

    createTable  db
    insertValues db
    selectValues db

 '
 '  Cleaning up
 '

    db.execute("drop table guids")

end sub ' }

sub createTable(db as dao.database) ' {

    db.execute(   _
   "create table guids ("        & _
   "  id     guid primary key, " & _
   "  txt    varchar(60)       " & _
   ")")


end sub ' }

sub insertValues(db as dao.database) ' {

  '
  '  Create queryDef for insert statement:
  '

     dim stmt    as dao.queryDef
     set stmt = db.createQueryDef("",   _
       "parameters "                     & _
       "  parid     varchar(38), "       & _
       "  parTxt    varchar(60); "       & _
       "insert into guids values ([parId], [parTxt]) ")


     dim valId   as dao.parameter
     dim valTxt  as dao.parameter

     set valId  = stmt.parameters("parId" )
     set valTxt = stmt.parameters("parTxt")

     dim g as guid

  '
  '  Insert ten records
  '

     dim i as long
     for i = 1 to 10

         g = CoCreateGuid_

         valId  = StringFromGUID2_(g)
         valTxt = StringFromGUID2_(g)

         stmt.execute
      next i


end sub ' }

sub selectValues(db as dao.database) ' {

    dim stmt as queryDef
    set stmt = db.createQueryDef("", "select * from guids")

    dim rs as dao.recordSet
    set rs = stmt.openRecordSet

    do while not rs.eof

       debug.print(rs("id") & "  " & rs("txt"))

       rs.moveNext
    loop

end sub ' }
