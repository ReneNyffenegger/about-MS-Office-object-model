option explicit

sub main()

    dim subName as string
    subName = "subTwo"

    application.run subName

end sub

sub subOne()
    msgBox "sub one"
end sub

sub subTwo()
    msgBox "sub two"
end sub

sub subThree()
    msgBox "sub three"
end sub
