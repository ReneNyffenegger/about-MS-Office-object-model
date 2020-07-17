option explicit

sub main() ' {


    dim sh as worksheet : set sh = activeSheet

    dim rng as range
    set rng = createTestData(sh)

    printFilterProperties sh              ' autoFilterMode: False, filterMode: False


  '
  ' Turn on the filters on the ranges.
  ' Note: the word «autoFilter» is a
  ' method when applied to a range,
  ' but a property when applied to
  ' a worksheet
  '
    rng.autoFilter
    printFilterProperties sh              ' autoFilterMode: True, filterMode: False



  '
  ' First criteria: value of colTwo
  ' needs to be between 50 and 100:
  '
    rng.autoFilter field     :=      2, _
                   criteria1 := ">50" , _
                   operator  := xlAnd , _
                   criteria2 := "<100"

    printFilterProperties sh              ' autoFilterMode: True, filterMode: True


  '
  ' Second critera: value of first column
  ' needs to start with the letter «U»:
  '
    rng.autofilter  field     :=  1,    _
                    criteria1 := "=U*"

    sh.usedRange.columns.autoFit

end sub ' }

function createTestData(sh as worksheet) as range ' {

'   sh.usedRange.clearFormats
'   sh.usedRange.clearContents

    dim r, c as long : r = 0 : c = 1

    with sh ' {

         r = r + 1 : .range( .cells(r, c), .cells(r, c+3) ).value = array("colOne", "colTwo", "colThree"  , "colFour")
         r = r + 1 : .range( .cells(r, c), .cells(r, c+3) ).value = array("AB"    ,     130 , #2014-06-02#,      21.4)
         r = r + 1 : .range( .cells(r, c), .cells(r, c+3) ).value = array("UVW"   ,      99 , #2010-05-07#,      17.2)
         r = r + 1 : .range( .cells(r, c), .cells(r, c+3) ).value = array("PQ"    ,      42 , #2020-02-17#,      38.0)
         r = r + 1 : .range( .cells(r, c), .cells(r, c+3) ).value = array("XYZ"   ,     111 , #2018-12-03#,       1.1)
         r = r + 1 : .range( .cells(r, c), .cells(r, c+3) ).value = array("DEFG"  ,      15 , #2017-04-28#,       5.9)
         r = r + 1 : .range( .cells(r, c), .cells(r, c+3) ).value = array("UTA"   ,     128 , #2011-06-03#,      33.9)
         r = r + 1 : .range( .cells(r, c), .cells(r, c+3) ).value = array("XYZ"   ,     111 , #2018-12-03#,       1.1)
         r = r + 1 : .range( .cells(r, c), .cells(r, c+3) ).value = array("CLM"   ,      68 , #2021-04-19#,      65.3)

         set createTestData = .range(.cells(1, c), .cells(r, c+3) )

    end with ' }

end function ' }

sub printFilterProperties(sh as worksheet) ' {

    debug.print "autoFilterMode: " & sh.autoFilterMode & ", filterMode: " & sh.filterMode

end sub ' }
