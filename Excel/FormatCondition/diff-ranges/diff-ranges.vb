option explicit

sub diff_ranges(         _
       ids_A   as range, _
       data_A  as range, _
       ids_B   as range, _
       data_B  as range )


   dim fc as formatCondition
   dim c as string : c = ","


 '
 ' Highlight keys in ids_A which are not found in ids_B and...
 '
   set fc = ids_A.formatConditions.add(xlExpression, formula1 := "=ISNA(MATCH(" & ids_A.cells(1,1).address(rowAbsolute := false) & c & ids_B.address & c & "0))")
   fc.interior.color = rgb(255,  40, 60)

 '
 ' keys in ids_B which are not found in ids_A:
 '
   set fc = ids_B.formatConditions.add(xlExpression, formula1 := "=ISNA(MATCH(" & ids_B.cells(1,1).address(rowAbsolute := false) & c & ids_A.address & c & "0))")
   fc.interior.color = rgb(255,  40, 60)

   dim formula as string

   formula  = "=" & data_A.cells(1,1).address(rowAbsolute := false, columnAbsolute := false)     & _
    " <> offset(" & data_B.cells(1,1).address(rowAbsolute := true , columnAbsolute := false) & c & _
    "match(" & ids_A.cells(1,1).address(rowAbsolute := false) & c & ids_B.address & ", 0)-1,0)"

   set fc = data_A.formatConditions.add(xlExpression, formula1 := formula)
   fc.interior.color = rgb(255, 217, 102)


   formula  = "=" & data_B.cells(1,1).address(rowAbsolute := false, columnAbsolute := false)     & _
    " <> offset(" & data_A.cells(1,1).address(rowAbsolute := true , columnAbsolute := false) & c & _
    "match(" & ids_B.cells(1,1).address(rowAbsolute := false) & c & ids_A.address & ", 0)-1,0)"

   set fc = data_B.formatConditions.add(xlExpression, formula1 := formula)
   fc.interior.color = rgb(255, 217, 102)

end sub
