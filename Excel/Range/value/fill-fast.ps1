# vi: ft=conf
$xls = new-object -com excel.application
$xls.visible = $true

$wb  = $xls.workbooks.add()
$sh  = $wb.worksheets.add()


(measure-command {
#
#  Uncommenting the following three statements
#  seem to slow down (rather than speed up)
#  the execution!
#
#  $xls.calculation    = -4135
#  $xls.screenUpdating =  $false
#  $xls.enableEvents   =  $false

   $nofRows = 100
   $nofCols =  50

   $values = new-object 'Object[,]' $nofRows, $nofCols

   foreach ($r in 0 .. ($nofRows-1) ) {
   foreach ($c in 0 .. ($nofCols-1) ) {
      $values[$r,$c] = "${r} * $c = $($r*$c)"
   }}

   $sh.range(
      $sh.cells(       1,       1),
      $sh.cells($nofRows,$nofCols)
   ).value = $values

#  $xls.screenUpdating =  $true
#  $xls.enableEvents   =  $true
#  $xls.calculation    = -4105
}).TotalSeconds

$xls = $null
