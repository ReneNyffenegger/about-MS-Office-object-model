option explicit

sub main() ' {

    dim rng as range
    set rng = [ r3c2:r5c6 ]

    rng.interior.color = rgb(255, 190, 60)

    cells(6,2) = "Dimensions before deleting: " & rng.rows.count & " x " & rng.columns.count

    columns(4).delete
    rows(4).delete

    cells(6,2) = "Dimensions after deleting: " & rng.rows.count & " x " & rng.columns.count

    rng.borderAround weight := xlThick, color := rgb(220, 140, 30)

end sub ' }
