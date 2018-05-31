private sub worksheet_change(byVal cell as range)
 '
 '  Check if user changed value in 2nd column:
    if cell.column = 2 then
     '
     ' Change cell's color one to the left
     '
       cell.offset(0, -1).interior.color = rgb(200, 130, 30)
    end if
    
end sub
