sub main()

   dim rng as range

   set rng = range("d7")

 '
 ' Resize range to seven rows and
 ' five columns columns:
 '
   set rng = rng.resize(7, 5)

 '
 ' Color the range in orange:
 '
   rng.interior.color = rgb(255, 127, 0)

 '
 ' Create another range based on the
 ' existing one. Apparently, it uses
 ' the one by one top left cell for the resize
 '
   set rng = rng.resize(4, 3)
   rng.interior.color = rgb(0, 45, 200)

end sub
