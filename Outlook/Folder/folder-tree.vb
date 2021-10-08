option explicit

sub main() ' {

    dim str as outlook.store

    for each str in application.session.stores
        dim rootFld as folder
        set rootFld = str.getRootFolder
        debug.print  str.filePath
        recurseFolder rootFld, 1

    next str

end sub ' }

sub recurseFolder(byVal fldParent as outlook.folder, byVal lvl as long) ' {

    debug.print string(lvl*2, " ") & fldParent.folderPath & " (" & fldParent.name & ")"

    dim fldChild as outlook.folder
    for each fldChild in fldParent.folders
        recurseFolder fldChild, lvl+1
    next fldChild

end sub ' }
