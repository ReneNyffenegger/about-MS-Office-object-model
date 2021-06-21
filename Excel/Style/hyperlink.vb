option explicit

sub main() ' {

    dim stHyperlink         as style
    dim stHyperlinkFollowed as style

    set stHyperlink         = thisWorkbook.styles.add("Hyperlink"         )
    set stHyperlinkfollowed = thisWorkbook.styles.add("Followed Hyperlink")

    stHyperlink.font.color         = rgb(5, 99, 193)
    stHyperlinkFollowed.font.color = rgb(5, 99, 193)

    activeSheet.cells(1,1).formula = "=hyperlink(""#b1"", ""b1"")"
    activeSheet.cells(2,1).formula = "=hyperlink(""#a1"", ""a1"")"

end sub ' }
