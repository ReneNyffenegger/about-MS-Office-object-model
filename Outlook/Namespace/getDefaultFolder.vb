option explicit


sub main() ' {

   printFolderPath  "olFolderCalendar"               , olFolderCalendar                '  9
   printFolderPath  "olFolderConflicts"              , olFolderConflicts               ' 19 (subfolder of the Sync Issues folder. Only available for an Exchange account)
   printFolderPath  "olFolderContacts"               , olFolderContacts                ' 10
   printFolderPath  "olFolderDeletedItems"           , olFolderDeletedItems            '  3
   printFolderPath  "olFolderDrafts"                 , olFolderDrafts                  ' 16
   printFolderPath  "olFolderInbox"                  , olFolderInbox                   '  6
   printFolderPath  "olFolderJournal"                , olFolderJournal                 ' 11
   printFolderPath  "olFolderJunk"                   , olFolderJunk                    ' 23
   printFolderPath  "olFolderLocalFailures"          , olFolderLocalFailures           ' 21 (subfolder of the Sync Issues folder - Only available for an Exchange account)
   printFolderPath  "olFolderManagedEmail"           , olFolderManagedEmail            ' 29
   printFolderPath  "olFolderNotes"                  , olFolderNotes                   ' 12
   printFolderPath  "olFolderOutbox"                 , olFolderOutbox                  '  4
   printFolderPath  "olFolderSentMail"               , olFolderSentMail                '  5
   printFolderPath  "olFolderServerFailures"         , olFolderServerFailures          ' 22 (subfolder of the Sync Issues folder. Only available for an Exchange account)
   printFolderPath  "olFolderSuggestedContacts"      , olFolderSuggestedContacts       ' 30
   printFolderPath  "olFolderSyncIssues"             , olFolderSyncIssues              ' 20 (Only available for an Exchange account.)
   printFolderPath  "olFolderTasks"                  , olFolderTasks                   ' 13
   printFolderPath  "olFolderToDo"                   , olFolderToDo                    ' 28
   printFolderPath  "olPublicFoldersAllPublicFolders", olPublicFoldersAllPublicFolders ' 18 (The All Public Folders folder in the Exchange Public Folders store - Only available for an Exchange account)
   printFolderPath  "olFolderRssFeeds"               , olFolderRssFeeds                ' 25

end sub '

sub printFolderPath(fldTxt as string, fldVal as olDefaultFolders) ' {
 on error goto err_
    dim fld as folder
    set fld = application.session.getDefaultFolder(fldVal)
    debug.print format(fldTxt & ": ", "!" & string(25, "@")) & fld.folderPath & " - " & fld.description
    exit sub
 err_:
    debug.print format(fldTxt & ": ", "!" & string(25, "@")) & "Error: " & err.description
end sub ' }
