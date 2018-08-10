sub main()

    range("b3:c5").select

  '
  ' After selecting a range, application.selection accordingly
  ' is a range. typeName(…) prints "Range".
  '
    debug.print typeName(application.selection)

  '
  ' We can now use application.selection to change
  ' the properties of the selected range
  '
    with application.selection
        .font.name      = "Courier New"
        .numberFormat   = "0.000"
        .interior.color = rgb(220,220,255)
    end with

end sub
