option explicit

sub main() ' {

    dim db as database
    set db = currentDb()

    dim t as tableDef
    set t = db.tableDefs("MSysObjects")

    dim col as field
    for each col in t.fields
        debug.print(col.name)
    next col

end sub ' }
