option explicit

sub main() ' {

   dim rng as range
   set rng = range(cells(4,3), cells(7, 7))
   rng.borderAround xlContinuous, xlMedium, color := rgb(140, 90, 180)

   setValidation rng

end sub ' }


sub setValidation(rng as range) ' {

    dim firstCellRelativeAddress as string
    dim formula                  as string

    firstCellRelativeAddress =  rng.address(rowAbsolute := false, columnAbsolute := false)
    formula                  = "=isNumber(" & firstCellRelativeAddress & ")"

    with rng.validation ' {

        .add type := xlValidateCustom, formula1 := formula

        .ignoreBlank  =  true

        .showInput    =  true
        .inputTitle   = "Validation rule"
        .inputMessage = "Enter a numerical value"

        .showError    =  true
        .errorTitle   = "Validation rule failed"
        .errorMessage = "Please enter a number""

    end with ' }

end sub ' }
