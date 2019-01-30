option explicit

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

  '
  ' Drop table if it exists:
  '
    if not isNull(dLookup("Name", "MSysObjects", "Name='tq84_tab'")) then db.execute("drop table tq84_tab" )

  '
  ' Create table and fill with one initial value:
  '
    db.execute("create table tq84_tab (col_1 number, col_2 varchar(100))")
    db.execute("insert into tq84_tab values(1, '*')")

    dim stmt as queryDef
    set stmt =  db.createQueryDef("", "insert into tq84_tab(col_1, col_2) select 2*col_1, col_1 & col_1 from tq84_tab")

    stmt.execute
    debug.print("Number of records inserteed: " & stmt.recordsAffected)

    stmt.execute
    debug.print("Number of records inserteed: " & stmt.recordsAffected)

    stmt.execute
    debug.print("Number of records inserteed: " & stmt.recordsAffected)

    stmt.execute
    debug.print("Number of records inserteed: " & stmt.recordsAffected)

end sub ' }
