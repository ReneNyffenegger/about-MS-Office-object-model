option explicit

sub main() ' {

    if not isNull(dLookup("Name", "MSysObjects", "Name='tq84_tab'")) then
       doCmd.close acTable, "tq84_tab", acSaveNo
       execSQL "drop table tq84_tab"
    end if

    execSQL "create table tq84_tab (num number, txt varchar(20))"

    execSQL "insert into tq84_tab values ( 1, 'one'   )"
    execSQL "insert into tq84_tab values ( 2, 'two'   )"
    execSQL "insert into tq84_tab values ( 3, 'three' )"
    execSQL "insert into tq84_tab values ( 4, 'four'  )"


    dim db as dao.database
    set db = currentDB()

    dim rs as dao.recordSet
    set rs = db.openRecordset("select * from tq84_tab order by txt", dao.dbOpenForwardOnly)

    do while not rs.eof ' {
       debug.print rs!num & "  " & rs!txt
       rs.moveNext
    loop ' }

    set rs = nothing

end sub ' }

sub execSQL(stmt as string) ' {
    currentProject.connection.execute(stmt)
end sub ' }
