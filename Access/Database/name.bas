option explicit

sub main()

    dim db as database
    set db = currentDb()

    debug.print(db.name)

end sub
