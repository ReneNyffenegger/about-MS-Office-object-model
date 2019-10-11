option explicit

sub main() ' {

    selection.pageSetup.orientation = wdOrientLandscape

    with selection
     '
     ' Make a really big font
     '
     .font.size = 120

     .typeText "This is page number "
     .fields.add selection.Range, text := "page"

     .typeParagraph

     .typeText "This is page number "
     .fields.add selection.Range, text := "page"

    end with

end sub ' }
