option explicit

sub main() ' {

    dim fold as folder
    set fold = application.session.getDefaultFolder(olFolderInbox)

    dim expl as explorer
    set expl = application.explorers.add(fold)

'   expl.display
    expl.activate

end sub ' }