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

    dim tlApp  as new tli.tliApplication
    dim tlInfo as     tli.typeInfo

    dim dummy  as     long

    set tlInfo = tlApp.interfaceInfoFromObject(obj)

    dim mbrInfo as tli.memberInfo
    for each mbrInfo in tlInfo.members ' {


        if mbrInfo.descKind <> tli.desckind_funcDesc then
           debug.print "Unexpected desckind for " & mbrInfo.name
'          goto skip
        end if

'       if mbrInfo.returnType.varType = tli.VT_EMPTY or mbrInfo.returnType.varType = tli.VT_NULL then
'          debug.print "Skiping " & mbrInfo.name
'          goto skip
'       end if

        if mbrInfo.name = "DefaultFilePath" or mbrInfo.name = "DefaultEPostageApp" or mbrInfo.name = "Entries" or mbrInfo.name = "FirstLetterExceptions" or mbrInfo.name = "TwoInitialCapsExceptions" or mbrInfo.name = "HangulAndAlphabetExceptions"  or mbrInfo.name = "OtherCorrectionsExceptions" then
'          debug.print "Skiping " & mbrInfo.name
           goto skip
        end if


        if      mbrInfo.invokeKind = invoke_propertyGet then

                dim settingValue as variant
                settingValue = callByName(obj, mbrInfo.name, vbGet)

'               debug.print mbrInfo.name

                if fillDict then
                   settings_.add mbrInfo.name, settingValue
                else

                  if settings_(mbrInfo.name) <> settingValue then
                     debug.print "Diff for " & mbrInfo.name & ": old = " & settings_(mbrInfo.name) & ", new = " & settingValue
                     settings_(mbrInfo.name) = settingValue
                  end if

                end if

        elseIf  mbrInfo.invokeKind = invoke_propertyPut then

              '
              ' Do nothing
              '

        else
'               debug.print "Unexpected invokeKind for " & mbrInfo.name
        end if

skip:

    next mbrInfo ' }

end sub ' }
