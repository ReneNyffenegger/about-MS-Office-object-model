option explicit

sub main() ' {

    dim interesting() as variant
        interesting = array("foo", "bar", "baz")
    
    if not isError(application.match("bla", interesting, 0)) then debug.Print "bla is interesting"
    if not isError(application.match("baz", interesting, 0)) then debug.Print "baz is interesting"

end sub ' }
