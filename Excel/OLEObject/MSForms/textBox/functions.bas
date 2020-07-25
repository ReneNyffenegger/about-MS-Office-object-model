option explicit

sub main() ' {

    with activeSheet ' {

         dim rng as range
         set rng = .range(.cells(2,2), .cells(12, 7))

         dim obj as oleObject
         set obj = activeSheet.OLEObjects.add( _
            classType       := "Forms.TextBox.1"  , _
            link            :=  false             , _
            displayAsIcon   :=  false             , _
            left            :=  rng.left          , _
            top             :=  rng.top           , _
            width           :=  rng.width         , _
            height          :=  rng.height          _
         )

    end with ' }

    dim tb as msForms.textBox
    set tb = obj.object

    with tb ' {

      .name              = "txt"
      .font              = "Courier"
      .enterKeyBehavior  =  true
      .multiLine         =  true
      .borderStyle       =  fmBorderStyleSingle
      .backColor         =  rgb(255, 240, 100)

    end with ' }

end sub ' }
