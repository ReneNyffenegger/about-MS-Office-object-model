option explicit

sub main() ' {

    iterateOverLinkSources xlExcelLinks , "xlExcelLinks"
    iterateOverLinkSources xlOLELinks   , "xlOLELinks"
    iterateOverLinkSources xlPublishers , "xlPublishers"   ' Macintosh only
    iterateOverLinkSources xlSubscribers, "xlSubscribers"  ' Macintosh only

end sub ' }

sub iterateOverLinkSources(tp as long, tpTxt as string) ' {

    dim lnkSrcs as variant
    dim lnkSrc  as variant

    lnkSrcs = activeWorkbook.linkSources(tp)

    if not isEmpty(lnkSrcs) then
       debug.print tpTxt
       for each lnkSrc in lnkSrcs
           debug.print "  " & lnkSrc
       next lnkSrc
    else
       debug.print tpTxt & " is empty"
    end if

end sub ' }
