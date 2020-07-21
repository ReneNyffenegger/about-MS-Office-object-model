option explicit

sub main() ' {

    dim ws_1, ws_2, ws_3 as worksheet

    set ws_1 = activeSheet    : ws_1.name = "foo"
    set ws_2 = worksheets.add : ws_2.name = "bar"
    set ws_3 = worksheets.add : ws_3.name = "baz"

    dim target_cell_1, target_cell_2, target_cell_3 as range

    set target_cell_1 = ws_1.cells( 3,  2)
    set target_cell_2 = ws_2.cells(90, 60)
    set target_cell_3 = ws_3.cells(60, 90)

    ws_1.hyperlinks.add anchor := target_cell_1, address := "", subAddress := target_cell_2.parent.name & "!" & target_cell_2.address, screenTip := "goto Target 2", textToDisplay := "target 2"
    ws_2.hyperlinks.add anchor := target_cell_2, address := "", subAddress := target_cell_3.parent.name & "!" & target_cell_3.address, screenTip := "goto Target 3", textToDisplay := "target 3"
    ws_3.hyperlinks.add anchor := target_cell_3, address := "", subAddress := target_cell_1.parent.name & "!" & target_cell_1.address, screenTip := "goto Target 1", textToDisplay := "target 1"

    ws_1.activate


end sub ' }
