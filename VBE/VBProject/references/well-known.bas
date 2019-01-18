' Microsoft Excel 15.0 Object Library
'     Excel
  call application.VBE.activeVBProject.references.addFromGuid("{00020813-0000-0000-C000-000000000046}", 1, 8)

' Microsoft VBScript Regular Expressions 5.5
'      VBScript_RegExp_55
'      C:\windows\SysWOW64\vbscript.dll\3
  call application.VBE.activeVBProject.references.addFromGuid("{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5, 5)

' Microsoft Scripting Runtime
'       Scripting
'       C:\Windows\SysWOW64\scrrun.dll         (or scrob.dll?)
  call application.VBE.activeVBProject.references.addFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)

' Windows Script Host
'      IWshRuntimeLibrary
'      C:\Windows\SysWOW64\wshom.ocx
  call application.VBE.activeVBProject.references.addFromGuid("{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}", 1, 0)

' Microsoft Visual Basic for Applications Extensibility
'       VBEIDE
'       C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
  call application.VBE.activeVBProject.references.addFromGuid("{0002E157-0000-0000-C000-000000000046}", 5, 3)

' Microsoft XML
'      MSXML6
'      C:\windows\System32\msxml6.dll
  call application.VBE.activeVBProject.references.addFromGuid("{F5078F18-C551-11D3-89B9-0000F81FE221}", 6, 0)

' WinHTTP
'     C:\windows\system32\winhttp.dll
  call application.VBE.activeVBProject.references.addFromGuid("{662901fc-6951-4854-9eb2-d9a2570f2b2e}", 5, 1)

' ADODB (Microsoft ActiveX Data Objects 6.1 Library)
'     C:\Program Files (x86)\Common Files\System\ado\msado15.dll
  call application.VBE.activeVBProject.references.addFromGuid("{B691E011-1797-432E-907A-4D8C69339129}", 6, 1)
