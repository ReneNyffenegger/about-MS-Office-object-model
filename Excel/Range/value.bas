option explicit

sub main()

  ' Define a range ...
  '
    dim rng as range
    set rng =  range("b2:e4")

  ' and set each cell's value within the
  ' range to the same value:
  '
    rng.value = "*"

  ' Define another range...
  '
    set rng = range("b6:d6")


  ' and set the values of each column in
  ' the range:
  '
    rng.value = array("foo", "bar", "baz")

end sub
