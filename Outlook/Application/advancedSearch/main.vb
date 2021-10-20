option explicit

sub main() ' {

    dim scope as string
        scope = "'" & session.getDefaultFolder(olFolderInbox   ).folderPath & "'," & _
                "'" & session.getDefaultFolder(olFolderSentMail).folderPath & "'"

    dim query as string
        query = "urn:schemas:mailheader:subject = 'XYZ'"

    advancedSearch scope, query, searchSubFolders := false, tag := "tq84"

end sub ' }
