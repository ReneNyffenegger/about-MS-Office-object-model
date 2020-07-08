option explicit

sub main() ' {

    dim days   as variant : days   = array("Mon", "Tue", "Wed", "Thu", "Fri")
    dim region as variant : region = array("South", "East", "West", "North" )

    cells(3,2).resize(5,1).value = worksheetFunction.transpose(days)
    cells(2,3).resize(1,4).value = region

end sub ' }
