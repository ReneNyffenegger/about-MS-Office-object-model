#if 0 then

thisWorkbook.VBProject.references.addFromGuid _
        GUID   :="{0002E157-0000-0000-C000-000000000046}", _
        major  :=  5,                                      _
        minor  :=  3

#end if

sub main() ' {

    dim comps as vbComponents
    set comps = application.vbe.activeVBProject.vbComponents

    dim comp  as vbComponent
    set comp  = comps.add(1)

    dim mdl as codeModule
    set mdl = comp.codeModule

    mdl.insertLines  1, "sub xyz(nm as long, tx as string)"
    mdl.insertLines  2, "   msgBox tx & "", the number is "" & nm"
    mdl.insertLines  3, "end sub"

    application.run "xyz", 42, "hello world"

    comps.remove comp

end sub ' }
