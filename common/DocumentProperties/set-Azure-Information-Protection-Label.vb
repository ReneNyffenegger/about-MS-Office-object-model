option explicit

sub mai() ' {

    dim wb as workbook
    set wb = workbooks.add

    dim guid_internal_use_only as string
        guid_internal_use_only = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

    debug.print typename(wb.customDocumentProperties)

    wb.customDocumentProperties.add "MSIP_Label_" & guid_internal_use_only & "_Enabled", false, msoPropertyTypeString, "true"

    wb.saveAs environ("temp") & "\created-at-" & format(now(), "yyyy-mm-dd_hh-nn")

    debug.print wb.name

end sub ' }
