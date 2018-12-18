option explicit

sub main() ' {

    dim colStart as long
    dim colEnd   as long

    colStart = 4
    colEnd   = 9

    range(columns(colStart), columns(colEnd)).select

    with selection
        .columnWidth    = application.centimetersToPoints(0.1)
        .interior.color = rgb(250, 140, 30)
    end with

end sub ' }
