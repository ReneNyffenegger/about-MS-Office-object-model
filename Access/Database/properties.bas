option explicit

sub main() ' {

    dim db as database
    set db = currentDb()

    dim prop as property
    for each prop in db.properties

        if prop.name <> "Connection" then
           debug.print(rpad(prop.name, 42) & " = " & prop.value)
        end if

    next prop

end sub ' }

'
' https://renenyffenegger.ch/notes/development/languages/VBA/modules/Common/Text
'
function rpad(text as String, length as integer, optional padChar as string = " ") ' {
    rpad = text & string(length - len(text), padChar)
end function ' }
