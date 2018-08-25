option explicit

sub main() ' {

    dim ary(1 to 5) as string

    ary(1) = "one"
    ary(2) = "two"
    ary(3) = "three"
    ary(4) = "four"
    ary(5) = "five"

    range("b3").resize(1, uBound(ary)).value = ary

end sub ' }
