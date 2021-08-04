option explicit

sub main() ' {

    dim rng as range

  '
  ' Create an 'original' range' ...
  '
    set rng = range(cells(3, 6), cells(14, 11))

  '
  ' ... and draw an orange border around it:
  '
    rng.borderAround xlContinuous, xlMedium, color := rgb(240, 110, 0)

  '
  ' Write the original range's coordinates (top left cell and dimensions) into
  ' its top left cell:
  '
    cells(rng.row, rng.column) = "orig: " & rng.columns.count & "x" & rng.rows.count & " @ " & rng.column & "," & rng.row

  '
  ' Create another range that is moved (off-set) by 3 rows below and 2 columns to the left
  ' of the original range, and draw a blue dashed around it:
  '
    dim rngOffset as range
    set rngOffset = rng.offset(3, -2)
    rngOffset.borderAround xlDash, xlMedium, color := rgb(100,  40,  200)

  '
  ' Create another range from the off-set range and resize it, then
  ' draw a green dash-dotted border around it:
  '
    dim rngResized as range
    set rngResized = rngOffset.resize(3, 4)
    rngResized.borderAround xlDashDot, xlMedium, color := rgb(50, 200, 80)

  '
  ' Resize column widths to make output less wide:
  '
    range(cells(1, 1), cells(1, 16)).columnWidth = 2

    cells(20, 20).select

end sub ' }
