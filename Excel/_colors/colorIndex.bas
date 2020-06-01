option explicit

const nofColorIndices = 56

sub main() ' {

    dim colorIndex as long
    for colorIndex = 1 to nofColorIndices ' {

        dim row as long
        dim col as long

        row = (colorIndex-1) \   8 + 1
        col = (colorIndex-1) mod 8 + 1

        with cells(row, col) ' {

             .value               =  colorIndex
             .interior.colorIndex =  colorIndex
             .font.colorIndex     = (colorIndex + 11) mod nofColorIndices

        end with ' }

    next colorIndex ' }

    range(columns(1), columns(8)).columnWidth = 6
    range(rows   (1), rows   (8)).rowHeight   = 38

    with range(cells(1,1), cells(nofColorIndices \ 8, 8))

        .font.size           = 20
        .font.bold           = true
        .horizontalAlignment = xlCenter
        .verticalAlignment   = xlCenter

    end with

    activeWorkbook.saved = true

end sub ' }
