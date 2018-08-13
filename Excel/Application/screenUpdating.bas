option explicit

sub main()

  Application.screenUpdating =  true
  call fill_10x40_rect(1, 1)
  
  Application.screenUpdating = false
  call fill_10x40_rect(1, 12)

end sub

sub fill_10x40_rect(startTopRow as long, startLeftColumn as long)

  dim t0, t1 as single
  dim r      as long
  dim c      as long
  
  t0 = timer
  for r = 0 to 39
  for c = 0 to  9
      cells(startTopRow + r, startLeftColumn + c).value = r * c
  next c
  next r
  t1 = timer
  
  debug.print "Filling 40 times 10 cells took " & (t1-t0) & " seconds."

end sub
