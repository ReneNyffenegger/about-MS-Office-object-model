option explicit

sub main()

    dim subName as string
    subName = "subTwo"

    application.run "runSub", "subTwo"

end sub

sub runSub(subName as string)
    debug.print "runSub tries to run " & subName
    application.run subName
end sub

sub subOne()
    debug.print "subOne was called"
end sub

sub subTwo()
    debug.print "subTwo was called"
end sub
