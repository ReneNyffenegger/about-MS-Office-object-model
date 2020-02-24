option explicit

sub main() ' {

    dim doc            as document
    dim fldNum, fldTxt as field

    set doc = documents.add

    doc.variables.add "num",  42
    doc.variables.add "txt", "Hello world"

    doc.content.insertAfter "Number = "

    set fldNum = doc.fields.add(endOfDocument(doc), wdFieldDocVariable, "num")

    doc.content.insertAfter chr(13)
    doc.content.insertAfter "Text = "
    set fldTxt = doc.fields.add(endOfDocument(doc), wdFieldDocVariable, "txt")

    fldNum.showCodes = false
    fldtxt.showCodes = false

end sub ' }

function endOfDocument(doc as document) as range ' {
  '
  ' Return the end of the document as a range.
  ' Not really sure if this is the most efficient and
  ' preferred way to do that.
  '

    set endOfDocument = doc.content
    endOfDocument.collapse wdCollapseEnd

end function ' }
