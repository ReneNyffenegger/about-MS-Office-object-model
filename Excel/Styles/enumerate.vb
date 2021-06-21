option explicit

sub main() ' {

    dim r  as long
    dim st as style

    for each st in thisWorkbook.styles ' {

        r = r + 1

        cells(r, 1)                = st.name

        cells(r, 2)                = st.font.name
        cells(r, 2).font.name      = st.font.name

        cells(r, 3)                = st.font.color
        cells(r, 3).font.color     = st.font.color

        cells(r, 4)                = st.interior.color
        cells(r, 4).interior.color = st.interior.color

    next st ' }

end sub ' }
