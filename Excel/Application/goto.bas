option explicit

sub main() ' {

    application.goto "r20c9" ' Must use R1C1 reference style
    activeCell = "I went to row 20, column 9"

    application.goto "r10c14"
    activeCell = "I went to row 10, column 14"

    dim rng as range
    set rng = cells(5, 3)
    application.goto rng
    activeCell = "I went to row 5, column 3 via a range"

end sub ' }
