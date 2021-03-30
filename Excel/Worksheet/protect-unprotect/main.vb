option explicit

const password  = "fooBarBaz42"
const sheetname = "protected sheet"

sub main() ' {
    create_workbook_with_protected_sheet

    dim wb as workbook
    dim sh as worksheet

    set wb = workbooks.open(wb_name)
    set sh = wb.sheets(sheetname)

    if try_password(sh, "password1") then
       msgBox "Unexpectedly, the wrong(?) password worked"
    end if

    if try_password(sh, password) then
       msgBox "Success: worksheet was unprotected"
    end if

end sub ' }

function wb_name() as string ' {
    wb_name  = environ$("temp") & "\protection-test.xlsx"
end function ' }

sub create_workbook_with_protected_sheet() ' {

    if dir(wb_name) <> "" then
       kill wb_name
    end if

    dim wb as workbook
    dim sh as worksheet

    set wb = workbooks.add

    set sh  = wb.worksheets.add
    sh.name = sheetname

    sh.cells(1,1) = 42
    sh.cells(2,1) ="Hello world"

    sh.protect password

    wb.saveAs                           _
       fileName   := wb_name          , _
       fileFormat := xlOpenXMLWorkbook

    wb.close

end sub ' }


function try_password(sh as worksheet, pw as string) as boolean ' {

    on error resume next

    sh.unprotect pw

    if err.number = 1004 then ' {
    '
    '  Sheet could not be unprotected
    '
       try_password = false
       exit function

    end if ' }

    try_password = true

end function ' }
