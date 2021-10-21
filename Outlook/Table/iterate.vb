option explicit

sub main() ' {

    dim fld as folder
    set fld = session.getDefaultFolder(olFolderInbox)

    dim tbl as table
    set tbl = fld.getTable("[subject] = 'Abschied'")

    tbl.sort("receivedTime")

    while not tbl.endOfTable

       dim r as row
       set r = tbl.getNextRow()

       dim entryId as string
       entryId = r.item("entryId")

       dim msg as mailItem
       set msg = session.getItemFromId(entryId)

       debug.print  entryId & ": " & r.item("LastModificationTime") & " - " & r.item("Subject") & " (" & msg.sender & ")"

    wend

end sub ' }
