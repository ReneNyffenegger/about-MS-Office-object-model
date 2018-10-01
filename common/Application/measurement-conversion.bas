option explicit

sub main() ' {

    debug.print "  1 inch:   " & application.inchesToPoints      (  1) & " points."
    debug.print "  1 cm:     " & application.centimetersToPoints (  1) & " points."

    if application.name = "Microsoft Word" then
    '
    '  Excel apparently lacks a few conversion methodes.
    '
       debug.print "  1 mm:     " & application.millimetersToPoints(  1) & " points."
       debug.print "  1 point:  " & application.pointsToCentimeters(  1) & " cm."
       debug.print "  1 point:  " & application.pointsToMillimeters(  1) & " mm."
       debug.print "180 points: " & application.pointsToInches     (180) & " inches."
       debug.print "100 points: " & application.pointsToPicas      (100) & " picas."
       debug.print " 25 picas:  " & application.picasToPoints      ( 25) & " points."
    end if

end sub  ' }
