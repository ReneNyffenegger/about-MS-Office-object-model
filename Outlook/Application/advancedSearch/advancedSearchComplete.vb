option explicit

private sub application_advancedSearchComplete(byVal found as search) ' {
  '
  ' Put this event into 'ThisOutlookSession' module
  '
    debug.print "search with tag " & found.tag & " completed"

  '
  ' dim tab as table
  ' set tab = found.getTable

    dim email as mailItem
    for each email in found.results
         email.display
    next email

end sub ' }
