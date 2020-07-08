option explicit

sub main() ' {
    
    cells(2,2) = "Price" : cells(2, 3) = "Sold"
    cells(3,2) =  1.08   : cells(3, 3) =  4
    cells(4,2) =  2.27   : cells(4, 3) =  1
    cells(5,2) =  0.97   : cells(5, 3) =  3
    cells(6,2) =  1.58   : cells(6, 3) =  6

    names.add "price", refersTo := range(cells(3,2), cells(6,2))
    names.add "sold" , refersTo := range(cells(3,3), cells(6,3))

  '
  ' Format header
  '
    with range(cells(2,2), cells(2,3))
        .font.bold      = true
        .interior.color = rgb(245, 230, 250)
        .horizontalAlignment = xlRight
    end  with


    cells(8,2) = "Total"
    cells(8,3).formulaArray = "= sum(price * sold)"
    cells(8,3).font.bold    =  true
    
    with union(range("price"), cells(8,3))
        .numberFormat = "0.00"
        .font.name    ="Courier New"
    end  with

end sub ' }
