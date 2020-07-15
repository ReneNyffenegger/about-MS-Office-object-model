option explicit

public sub main()

    cells(1,1).value = "=3+4"
    cells(2,1).value = 0
    cells(3,1).value = 1

    with application.activeWindow

  ' Make formula visible ...
   .displayFormulas            = true

  ' ... but don't show formula bar.
  '    (Note the application. here)
  '
    application.displayFormulaBar = false

   .displayGridlines           = false

  ' let the column names (A ... ) and row numbers (1 ...)
  ' disappear
   .displayHeadings            = false

  ' no scrollbars
   .displayHorizontalScrollbar = false
   .displayVerticalScrollbar   = false

  ' If display ruler is false, the horizontal and
  ' vertical rulers won't be displayed, irrespective
  ' of their value
   .displayRuler               = false

    range( cells(2,1), cells(5,1) ).rows.group
  ' dont show grouping symbols (aka «outline»).
   .displayOutline             = false

   .displayRightToLeft         = false

   .displayWorkbookTabs        = false

   .displayZeros               = false

    end with

    application.width          = 200
    application.height         = 160

end sub
