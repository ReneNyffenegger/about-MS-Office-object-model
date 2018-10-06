option explicit

sub main() ' {

    dim obj as object
    set obj = cells(1, 1)

    getParent obj
    getParent obj
    getParent obj


end sub ' }

sub getParent(byRef obj as object) ' {

  '
  ' Save old typeName
  '
    dim typeName_ as string : typeName_ = typeName(obj)

  '
  ' get obj's Parent
  '
    set obj = obj.parent

  '
  ' print both old and new typename
  '
    debug.print typeName_ & " -> " &  typeName(obj)

end sub ' }
