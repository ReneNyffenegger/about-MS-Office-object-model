option explicit

sub main() ' {

    dim num as double
    dim txt as string
    dim bol as boolean

    num = application.inputBox("Enter a number" , "Example", type := 1) : debug.print("num = " & num)
    txt = application.inputBox("Enter a string" , "Example", type := 2) : debug.print("txt = " & txt)
    bol = application.inputBox("Enter a boolean", "Example", type := 4) : debug.print("bol = " & bol)

end sub ' }
