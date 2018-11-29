option explicit

sub main()

   dim doc as document
   dim pag as page

   set doc = activeDocument
   set pag = doc.pages(1)

   dim rect1 As Shape
   dim rect2 As Shape

   set rect1 = pag.drawRectangle(3, 4, 5, 3)
   set rect2 = pag.drawRectangle(3, 2, 5, 1)

   rect1.text = "One"
   rect2.text = "Two"

   dim dynConn as shape

 ' Set dynConn = doc.Pages(1).Shapes("Dynamic connector")

 '
 ' Connect the shapes rect1 and rect2.
 '
 ' The third parameter is required although I fail to see
 ' what it is needed for.
 '
   rect1.autoConnect rect2, visAutoConnectDirNone, dynConn

 '
 ' Unfortunately, AutoConnect does not return the shape (connector)
 ' it created.
 ' Thus, the most recent created shape is assumed to be the
 ' one with the highest count number:
 '
   set dynConn = pag.shapes(pag.shapes.count)

 '
 ' Change end of connect to an arrow:
 '
   dynConn.Cells("EndArrow") = 13

 '
 ' Don't show grid:
 '
   activeWindow.showGrid = false

end sub
