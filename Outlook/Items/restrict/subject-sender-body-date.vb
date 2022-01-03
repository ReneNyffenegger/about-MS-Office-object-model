option explicit

sub src()

    Dim map    as outlook.NameSpace
    Dim fld    as outlook.folder
    dim fnd    as outlook.items
    dim itm    as object
    dim filter as string

    Set map   = application.getNamespace("MAPI")
    Set fld   = map.getDefaultFolder(olFolderInbox)

    filter =             "@SQL="
    filter = filter &     " urn:schemas:httpmail:subject          like '%excel%'"
    filter = filter & " and urn:schemas:httpmail:sendername       like '%mayer%'"
    filter = filter & " and urn:schemas:httpmail:textdescription  like '%kindly remind%'"
    filter = filter & " and urn:schemas:httpmail:date > ' " & format$(#2021-10-02#, "general date") & "'"

    set fnd = fld.items.restrict(filter)

    debug.print "found " & fnd.count & " items"

    for each itm in fnd
        debug.print
        debug.print itm.subject
        debug.print "  From: " & itm.sender.name
        debug.print "  Sent: " & itm.creationTime
    next

end sub
