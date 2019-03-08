'
' Add reference for »Microsoft Scripting Runime« (https://renenyffenegger.ch/notes/development/languages/VBA/Useful-object-libraries/Microsoft-Scripting-Runtime)
' in order to have the dictionary:
'
'     thisDocument.VBProject.references.addFromGuid GUID:="{420B2830-E718-11CF-893D-00A0C9054228}", major := 1, minor := 0
'
' And also for Type Lib Information:
'
'     call application.VBE.activeVBProject.references.addFromGuid("{8B217740-717D-11CE-AB5B-D41203C10000}", 1, 0)
'
' ---------------------------------------------------------------------------------------------------------------
'
'    https://renenyffenegger.ch/notes/Microsoft/Office/Object-Model/scripts/index
'    https://github.com/ReneNyffenegger/about-MS-Office-object-model/blob/master/scripts/diff-object-properties.bas
'

option explicit

dim settings_ as scripting.dictionary
dim obj       as object

sub gatherProperties(o as object) ' {

    set settings_ = new scripting.dictionary
    set obj       = o

    iterateOverOptions true

end sub ' }

sub diffProperties() ' {

    iterateOverOptions false

end sub ' }

sub iterateOverOptions(fillDict as boolean) ' {

    on error goto err_

    dim tlApp  as new tli.tliApplication
    dim tlInfo as     tli.typeInfo

    dim dummy  as     long

    set tlInfo = tlApp.interfaceInfoFromObject(obj)

    dim mbrInfo as tli.memberInfo
    for each mbrInfo in tlInfo.members ' {

        if mbrInfo.descKind <> tli.desckind_funcDesc then ' {
           debug.print "Unexpected desckind for " & mbrInfo.name
'          goto skip
        end if ' }


'       if mbrInfo.returnType = tli.vt_empty then ' {
'          debug.print "Return type empty for " & mbrInfo.name
'          goto skip
'       end if '


        if hasMandatoryParameters(mbrInfo) then ' {
'          debug.print mbrInfo.name " has mandatory parameters, skipping"
           goto skip
        end if ' }

        if      mbrInfo.invokeKind = invoke_propertyGet then ' {

                dim settingValue as variant
                settingValue = callByName(obj, mbrInfo.name, vbGet)

                if fillDict then
                   settings_.add mbrInfo.name, settingValue
                else

                  if settings_(mbrInfo.name) <> settingValue then
                     debug.print "Diff for " & mbrInfo.name & ": old = " & settings_(mbrInfo.name) & ", new = " & settingValue
                     settings_(mbrInfo.name) = settingValue
                  end if

                end if

'       elseIf  mbrInfo.invokeKind = invoke_propertyPut then
        ' }
        else ' {
             '
             ' Do nothing
             '
        end if' ' }

skip:

    next mbrInfo ' }

    exit sub

err_:

    if err.number = 5852 then                  ' Application-defined or object-defined error
'      debug.print "5852 for " & mbrInfo.name  '   This error occurs for example when trying to access
       resume next                             '   application.options.defaultEPostageApp in MS Word
    end if                                     '

    if err.number =  450 then                  ' Wrong number of arguments or invalid property assignment
'      debug.print "450 for "  & mbrInfo.name  '   This error occurs for example when trying to access
       resume next                             '   application.documents in MS Word
    end if                                     '

    if err.number =  509 then                  ' The  command is not available because the command is not available on this platform.
'      debug.print "509 for "  & mbrInfo.name  '   This error occurs for example when trying to access
       resume next                             '   application.wordBasic in MS Word
    end if                                     '

    if err.number =  438 then                  ' Object doesn't support this property or method
'      debug.print "438 for "  & mbrInfo.name  '   This error occurs for example when trying to access
       resume next                             '   application.system in MS Word
    end if                                     '

    if err.number = 5111 then                  ' Application-defined or object-defined error
'      debug.print "5111 for " & mbrInfo.name  '   This error occurs for example when trying to access
       resume next                             '   application.fileSearch in MS Word
    end if                                     '

    if err.number =   91 then                  ' Object variable or With block variable not set
'      debug.print "  91 for " & mbrInfo.name  '   This error occurs for example when trying to access
       resume next                             '   application.hangulHanjaDictionaries in MS Word
    end if                                     '

    if err.number =  440 then                  ' Automation error
'      debug.print " 440 for " & mbrInfo.name  '   This error occurs for example when trying to access
       resume next                             '   application.answerWizard in MS Word
    end if                                     '

    debug.print "Error: " & err.description & " (" & err.number & ")"
    debug.print "       " & mbrInfo.name


end sub ' }

function hasMandatoryParameters(mbr as tli.memberInfo) as boolean

'     hasMandatoryParameters = false

      hasMandatoryParameters = (mbr.parameters.count <> mbr.parameters.defaultCount)

'     dim par as tli.parameterInfo
'     for each par in mbr.parameters ' {

'         if not (par.flags and tli.paramFlag_fOpt) then
'            hasMandatoryParameters = true
'            exit function
'         end if

'     next par ' }

end function ' }
