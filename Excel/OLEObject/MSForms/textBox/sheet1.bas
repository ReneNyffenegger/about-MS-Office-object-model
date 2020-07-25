option explicit

private sub txt_keyDown(byVal keyCode as msForms.returnInteger, byVal state as integer) ' {

    dim shift, alt, ctrl as string

    if state and 1 then shift = "shf" 
    if state and 2 then ctrl  = "ctl"
    if state and 4 then alt   = "alt"

    debug.print _
      shift    & chr(9) & _
      ctrl     & chr(9) & _
      alt      & chr(9) & _
      keyCode  & chr(9)

end sub ' }
