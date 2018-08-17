option explicit

sub main()

  fillData

  dim c as range

  for each c in range("a1:b5")
      debug.print "Value of " & c.address & " = " & c.value
  next c

end sub

sub fillData()

    cells(1, 1) = 17
    cells(2, 1) =  9
    cells(3, 1) = 48
    cells(4, 1) =  6
    cells(5, 1) = 33

    cells(1, 2) = 21
    cells(2, 2) = 14
    cells(3, 2) =  9
    cells(4, 2) = 27
    cells(5, 2) = 39

end sub
