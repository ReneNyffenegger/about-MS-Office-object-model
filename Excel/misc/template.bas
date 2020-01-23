option explicit

sub main() ' {

    dim xls as excel.application
    dim bok as excel.workbook
    dim sht as excel.workSheet

    set xls = new excel.application

    xls.visible = true

    set bok = xls.workbooks.add
    set sht = bok.worksheets.add

    sht.cells(1, 1) = "Hello world"
    sht.cells(2, 1) = "The number is:"
    sht.cells(2, 2) =  42

    sht.columns(1).autofit
    sht.columns(2).autofit

    dim fileName as string
    fileName = environ("TEMP") & "\" & "bla.xlsx"

  '
  ' Delete Excel workbook if already created
  ' in a previous run.
  '
    if dir(fileName) <> "" then kill fileName

    bok.saveAs fileName

end sub ' }
