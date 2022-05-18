option explicit

sub src()

    dim map    as outlook.NameSpace
    dim fld    as outlook.folder
    dim fnd    as outlook.items
    dim itm    as object
    dim filter as string

    Set map   = application.getNamespace("MAPI")
    Set fld   = map.getDefaultFolder(olFolderInbox)

    filter = "@SQL="
    filter = filter &     " urn:schemas:httpmail:subject          like '%digicert%'"
    filter = filter & " and urn:schemas:httpmail:sendername       like '%mayer%'"
    filter = filter & " and urn:schemas:httpmail:textdescription  like '%kindly remind%'"
    filter = filter & " and urn:schemas:httpmail:date > ' " & format$(#2021-10-02#, "general date") & "'"

    set fnd = fld.items.restrict(filter)

    debug.print "found " & fnd.count & " items"

    for each itm in fnd

        if typeName(itm) <>  "MailItem" then
         ' Only interested in MailItem, skip items auch as MeetingItem etc.
           goto skip_
        end if

        if itm.attachments.count = 0 then
         ' Print information only if mail has an attachment
           goto skip_
        end if
        
        debug.print
        debug.print itm.subject
        debug.print "  From    : " & itm.sender.name
        debug.print "  Sent    : " & itm.creationTime
        debug.print "  Entry ID: " & itm.entryId

      ' Show attachement names if there are any
        dim att as attachment
        for each att in itm.attachments
            debug.print "      " & att.fileName
        next att

        skip_:

    next

end sub
