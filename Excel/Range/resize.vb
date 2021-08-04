option explicit

sub main()

   dim rng as range

   set rng = range("b3")

 '
 ' Resize range to five rows and
 ' four columns columns:
 '
   set rng = rng.resize(5, 4)

 '
 ' Color the range in orange:
 '
   rng.interior.color = rgb(255, 127, 0)

 '
 ' Create another range based on the
 ' existing one. Apparently, it uses
 ' the one by one top left cell for the resize
 '
   set rng = rng.resize(3, 2)

 '
 ' Color the «new» range in blue:
 '
   rng.interior.color = rgb(0, 45, 200)

end sub
