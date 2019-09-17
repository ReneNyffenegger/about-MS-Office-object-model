option explicit

sub worksheet_selectionChange(byVal rng as range) ' {

    if rng.count <> 1 then
       exit sub
    end if

    if rng.row = 5 and rng.column >= 2 and rng.column <= 37 then
       dim cel as range
       for each cel in range(cells(5,2), cells(5,37)) 
           cel.comment.visible = false
       next cel

       rng.comment.visible = true
    end if

end sub ' }
