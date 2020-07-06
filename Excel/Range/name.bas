option explicit

public sub main() ' {

    range("b3:d6").name = "randomNumbers"
    range("e2:f8").name = "foos"

    range("randomNumbers").formula = "=rand()"
    range("randomNumbers").interior.color = rgb(255, 140,  30)

    range("foos").value   = "foo"
    range("foos").interior.color = rgb(145, 175, 255)

    activeWorkbook.saved = true

end sub ' }
