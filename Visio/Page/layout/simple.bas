option explicit

dim pag as page

sub main()

   dim doc as document

   set doc = activeDocument
   set pag = doc.pages(1)

   dim nodes(1 to 10) as shape
   dim edges(1 to 10) as shape

   set nodes( 1) = createNode("one"  )
   set nodes( 2) = createNode("two"  )
   set nodes( 3) = createNode("three")
   set nodes( 4) = createNode("four" )
   set nodes( 5) = createNode("five" )
   set nodes( 6) = createNode("six"  )
   set nodes( 7) = createNode("seven")
   set nodes( 8) = createNode("eight")
   set nodes( 9) = createNode("nine" )
   set nodes(10) = createNode("ten"  )

   set edges( 1) = createEdge(nodes( 3), nodes( 5))
   set edges( 2) = createEdge(nodes( 3), nodes( 1))
   set edges( 3) = createEdge(nodes( 4), nodes( 3))
   set edges( 4) = createEdge(nodes( 2), nodes( 8))
   set edges( 5) = createEdge(nodes( 3), nodes( 2))
   set edges( 6) = createEdge(nodes( 7), nodes( 3))
   set edges( 7) = createEdge(nodes( 5), nodes( 1))
   set edges( 8) = createEdge(nodes( 5), nodes( 9))
   set edges( 9) = createEdge(nodes( 1), nodes( 6))
   set edges(10) = createEdge(nodes( 8), nodes(10))

 '
 ' Automatically layout shapes on the page:
 '
   pag.layout

   activeWindow.showGrid = false

end sub

function createNode(txt as string) as shape ' {

    set createNode = pag.drawRectangle(0, 0, 1.2, 0.5)
  createNode.text = txt

end function ' }

function createEdge(nodeFrom as shape, nodeTo as shape) as shape ' {

    nodeFrom.autoConnect nodeTo, visAutoConnectDirNone, createEdge
    set createEdge = pag.shapes(pag.shapes.count)

  ' Add arrow
    createEdge.cells("EndArrow") = 13

end function ' }
