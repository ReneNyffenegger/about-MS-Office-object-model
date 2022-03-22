option explicit

sub test_diff_ranges(optional multiple_worksheets as boolean = false) ' {

    dim sh_1 as worksheet
    dim sh_2 as worksheet

    set sh_1 = worksheets.add

    if multiple_worksheets then
       set sh_2 = worksheets.add
    else
       set sh_2 = sh_1
    end if

    with sh_1
   .range(.cells(3, 3), .cells(3, 8)).value = array("ID"       , "val A", "val B", "val C", "val D", "val E")
   .range(.cells(4, 3), .cells(4, 8)).value = array("20-a-13"  , "a"    , "b"    , "c"    , "dd"   , "f"    )
   .range(.cells(5, 3), .cells(5, 8)).value = array("9-xy-8"   , "g"    , "hhh"  , "i"    , "by"   , "dx"   )
   .range(.cells(6, 3), .cells(6, 8)).value = array("28-uyy-3" , "l"    , "m"    , "n"    , "o"    , "p"    )
   .range(.cells(7, 3), .cells(7, 8)).value = array("7-hpp-9"  , "q"    , "r"    , "s"    , "t"    , "u"    )
   .range(.cells(8, 3), .cells(8, 8)).value = array("13-gv-2"  , "v"    , "w"    , "x"    , "y"    , "z"    )
    end with


    with sh_2
   .range(.cells(3,11), .cells(3,16)).value = array("ID"       , "val A", "val B", "val C", "val D", "val E")
   .range(.cells(4,11), .cells(4,16)).value = array("20-a-13"  , "a"    , "b"    , "c"    , "d"    , "f"    )
   .range(.cells(5,11), .cells(5,16)).value = array("7-hpp-9"  , "q"    , "r"    , "s"    , "j"    , "k"    )
   .range(.cells(6,11), .cells(6,16)).value = array("13-gv-2"  , "vvv"  , "w"    , "x"    , "y"    , "z"    )
   .range(.cells(7,11), .cells(7,16)).value = array("21-aed-72", "q"    , "r"    , "imi"  , "t"    , "u"    )
   .range(.cells(8,11), .cells(8,16)).value = array("13-uxd-8" , "y"    , "tr"   , "ul"   , "j"    , "k"    )
   .range(.cells(9,11), .cells(9,16)).value = array("9-xy-8"   , "g"    , "h"    , "i"    , "by"   , "dx"   )
    end with

    diff_ranges                         _
       sh_1.range(sh_1.cells(4,  3), sh_1.cells(8, 3)), _
       sh_1.range(sh_1.cells(4,  4), sh_1.cells(8, 8)), _
       sh_2.range(sh_2.cells(4, 11), sh_2.cells(9,11)), _
       sh_2.range(sh_2.cells(4, 12), sh_2.cells(9,16))

end sub ' }
