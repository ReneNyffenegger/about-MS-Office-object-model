'
'
'
option explicit

sub main() ' {

   rows(2).rowHeight = 22
   rows(4).rowHeight = 22

   addButton cells(2,2), "one", "action_1"
   addButton cells(4,2), "two", "action_2"

end sub ' }

sub addButton( _
                rng                as range , _
                caption            as string, _
                action             as string)


    dim btn as oleObject
    set btn = rng.parent.oleObjects.add( _
         classType     := "Forms.CommandButton.1"  , _
         link          :=  false                   , _
         displayAsIcon :=  false                   , _
         left          :=  rng.left                , _
         top           :=  rng.top                 , _
         width         :=  rng.width               , _
         height        :=  rng.height )

    btn.object.caption = caption

    dim line as long
    with thisWorkbook.vbProject.vbComponents(rng.parent.codeName).codeModule ' {
         line = .countOfLines
        .insertLines line + 1, "sub " & btn.name & "_click()"
        .insertLines line + 2, "  " & action
        .insertLines line + 3, "end sub"
    end with ' }


end sub ' }

sub action_1() ' {
    msgBox "action 1"
end sub ' }

sub action_2() ' {
    msgBox "action 2"
end sub ' }
