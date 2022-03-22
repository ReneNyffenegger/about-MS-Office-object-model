option explicit

sub diff_ranges(         _
       ids_A   as range, _
       data_A  as range, _
       ids_B   as range, _
       data_B  as range )


   dim fc as formatCondition
   dim c  as string : c = ","

 '
 ' Formulae can be shortened by setting paramter 'external' to false, but then, they
 ' will only work if the data being compared is on same sheet.
 ' For convenience, I always set this value to true, although it probably
 ' would be more elegant if 'external' is only used if so required.
 '
   dim multipleSheets as boolean : multipleSheets = true


 '
 ' Highlight keys in ids_A which are not found in ids_B and...
 '
   set fc = ids_A.formatConditions.add(xlExpression, formula1 := "=ISNA(MATCH(" & ids_A.cells(1,1).address(rowAbsolute := false, external := multipleSheets) & c & ids_B.address(external := multipleSheets) & c & "0))")
   fc.interior.color = rgb(255,  40, 60)

 '
 ' keys in ids_B which are not found in ids_A:
 '
   set fc = ids_B.formatConditions.add(xlExpression, formula1 := "=ISNA(MATCH(" & ids_B.cells(1,1).address(rowAbsolute := false, external := multipleSheets) & c & ids_A.address(external := multipleSheets) & c & "0))")
   fc.interior.color = rgb(255,  40, 60)

   dim formula as string

   formula  = "=" & data_A.cells(1,1).address(rowAbsolute := false, columnAbsolute := false, external := multipleSheets)     & _
    " <> offset(" & data_B.cells(1,1).address(rowAbsolute := true , columnAbsolute := false, external := multipleSheets) & c & _
    "match(" & ids_A.cells(1,1).address(rowAbsolute := false, external := multipleSheets) & c & ids_B.address(external := multipleSheets) & ", 0)-1,0)"

   set fc = data_A.formatConditions.add(xlExpression, formula1 := formula)
   fc.interior.color = rgb(255, 217, 102)


   formula  = "=" & data_B.cells(1,1).address(rowAbsolute := false, columnAbsolute := false, external := multipleSheets)     & _
    " <> offset(" & data_A.cells(1,1).address(rowAbsolute := true , columnAbsolute := false, external := multipleSheets) & c & _
    "match(" & ids_B.cells(1,1).address(rowAbsolute := false, external := multipleSheets) & c & ids_A.address(external := multipleSheets) & ", 0)-1,0)"

   set fc = data_B.formatConditions.add(xlExpression, formula1 := formula)
   fc.interior.color = rgb(255, 217, 102)

end sub
