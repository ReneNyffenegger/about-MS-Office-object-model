option explicit

sub main() ' {

    dim wb       as workbook : set wb = activeWorkbook
    dim linkType as long     : linkType = xlExcelLinks
    dim links    as variant  : links    = wb.linkSources(linkType)

    dim link as variant

    for each link in links

        wb.breakLink link, linkType

    next link


end sub ' }
