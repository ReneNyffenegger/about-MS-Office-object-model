option explicit

sub main() ' {

  '
  ' Create a new worksheet on which the
  ' hyperlinks will be created:
  '
    dim sht as worksheet
    set sht = worksheets.add

    with sht

     '
     '  Add three hyperlinks to the sheet
     '
       .hyperlinks.add anchor := .cells(1, 1), textToDisplay := "first link" , screenTip := "foo", address := ""
       .hyperlinks.add anchor := .cells(2, 1), textToDisplay := "second link", screenTip := "bar", address := ""
       .hyperlinks.add anchor := .cells(3, 1), textToDisplay := "third link" , screenTip := "baz", address := ""

       .usedRange.columns.autoFit

      '
      ' Get the worksheet's code module using the worksheet's codeName property:
      '
        dim cdm as vbide.codeModule
        set cdm = application.vbe.activeVBProject.vbComponents(.codeName).codeModule

    end with

  '
  ' Insert the code with the event handler into the code module:
  '
    cdm.insertLines 1, "option explicit"

    cdm.insertLines 2, "private sub worksheet_followHyperlink(byVal hLnk as hyperlink)"
    cdm.insertLines 3, "   msgBox ""Hyperlink was clicked, screenTip is: "" & hLnk.screenTip"
    cdm.insertLines 4, "end sub"

end sub ' }
