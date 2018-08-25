option explicit

sub main() ' {

    dim ary_2d as variant

    ary_2d = [{ 1,2,3,4,5 ; "one","two","three","four","five" ; "a","b","c","d","e" }]

    dim ary_width  as integer
    dim ary_height as integer

    ary_width  = uBound(ary_2d, 2)
    ary_height = uBound(ary_2d, 1)

    debug.print ary_width  ' 5
    debug.print ary_height ' 3

    range("b2").resize(ary_height, ary_width).value = ary_2d

end sub ' }
