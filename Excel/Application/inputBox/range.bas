option explicit

sub main() ' {

    dim rng as range

    set rng = application.inputBox("Enter a range"  , "Example", type := 8)

    debug.print("The entered range is: " & rng.address)

end sub ' }
