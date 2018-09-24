option explicit

sub main() ' {

    dim t as template
    for each t in templates
        debug.print t.name & " (" & t.path & ")"
    next t

end sub ' }
