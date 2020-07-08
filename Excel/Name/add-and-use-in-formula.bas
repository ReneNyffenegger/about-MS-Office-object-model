option explicit

sub main() ' {

    dim nameFruits, nameNumbers as name

    cells(1,1) = "Apple"
    cells(2,1) = "Banana"
    cells(3,1) = "Coconut"
    cells(4,1) = "Durian"
    cells(5,1) = "Elderberry"

    set nameFruits  = names.add("fruits",  refersTo := range(cells(1,1), cells(5,1))        )
    set nameNumbers = names.add("numbers", refersTo := array("one", "two", "three", "four") )

    cells(2, 3).formula = "= ""Third element in fruits is ""  & index(fruits , 3)"
    cells(3, 3).formula = "= ""Third element in numbers is "" & index(numbers, 3)"

end sub ' }
