option explicit

sub main() ' {

    dim ws as worksheet
    
    set ws = activeWorkbook.worksheets.add
    
    ws.cells(1, 1).value        =  43870.51
    ws.cells(1, 2).formulaR1C1  = "= RC[-1]"
    ws.cells(2, 1).value        =  43970.920347
    ws.cells(2, 2).value        = #2020-05-19 22:05:18#

    ws.cells(1, 2).numberFormat = "yyyy-mm-dd hh:mm:ss"


    checkForMisleadingFormat ws

end sub ' }

sub checkForMisleadingFormat(ws as worksheet) ' {

     dim c as range
     for each c in ws.usedRange ' {

         if c.errors(xlMisleadingFormat).value then ' {

            debug.print("Misleading format error found in " & c.address)

         end if' }

     next c ' }

end sub ' }
