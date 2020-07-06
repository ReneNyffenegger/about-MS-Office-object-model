option explicit

sub main() ' {

    testData

    searchInSecondColumn

    findWorld false
    findWorld true
 

end sub ' }

sub searchInSecondColumn() ' {

    dim found as range

    set found  = columns(2).find(            _
                    what   := "baz"        , _
                    after  :=  cells(1, 2) , _
                    lookIn :=  xlValues    , _
                    lookAt :=  xlWhole       _
                 )


    if found is nothing then
       debug.print("baz was not found in 2nd column")
    else
       debug.print("baz was found in 2nd column in row " & found.row)
    end if

end sub ' }

sub findWorld(whole as boolean) ' {

    debug.print("Searching for world, whole = " & whole)

    dim found as range

    dim wholeVal as integer
    if whole then
       wholeVal = xlWhole
    else
       wholeVal = xlPart
    end if

  '
  ' Search on the entire sheet.
  '
  ' If the searched value is in the cell that the
  ' after parameter specifies, it's returned last
  ' in the searching-cylce. Another stupidity, imho.
  '
    set found = cells().find (                    _
                    what        := "world"      , _
                    lookIn      :=  xlValues    , _
                    after       :=  cells(1, 1) , _
                    searchOrder :=  xlByColumns , _
                    lookAt      :=  wholeVal      _
                 )

    if found is nothing then
       debug.print("world was not found.")
       exit sub
    end if


  '
  ' This is soooo stupid. By default, calling findNext() wraps
  ' around. Thus, in order to prevent an infinite loop, we
  ' need to store the first address found and compare it
  ' to the subsequent addresses that are found to break the
  ' loop.
  '
    dim firstAddressFound as string
    firstAddressFound = found.address

    do ' {
       debug.print("found world in " & found.value & " at address " & found.address)

       set found = cells().findNext(after := found)

'      if found.address = firstAddressFound then exit do
'      msgBox "found world in " & found.value & " at address " & found.address

    loop until found is nothing or found.address = firstAddressFound ' }

end sub ' }

sub testData() ' {

    activeSheet.usedRange.clearContents

    cells(1, 1) = "Hello world" : cells(1, 2) = "Good bye" 
    cells(2, 1) = "foo"         : cells(2, 2) = "Another world was found on Mars"
    cells(3, 1) = "bar"         : cells(3, 2) = "baz"         
    cells(4, 1) = "baz"         : cells(4, 2) = "four"         
    cells(5, 1) = "five"        : cells(5, 2) = "bar"        
    cells(6, 1) = "foo"         : cells(6, 2) = "six"
    cells(7, 1) = "seven etc."  : cells(7, 2) = "foo"

    columns(1).autoFit
    columns(2).autoFit

end sub ' }
