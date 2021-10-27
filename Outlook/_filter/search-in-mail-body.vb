option explicit

sub searchInMailBodies(searchText as string) ' {

   dim fld as folder
   set fld = session.getDefaultFolder(olFolderInbox)

   dim DASL_query as string
   DASL_query = "@SQL=""urn:schemas:httpmail:textdescription"" ci_phrasematch '" & searchText & "'"

   dim tbl as table
   set tbl = fld.getTable(DASL_query)

   dim r as row
   do until tbl.endOfTable
      set r = tbl.GetNextRow
      debug.print r("Subject")
   loop

end sub ' }
