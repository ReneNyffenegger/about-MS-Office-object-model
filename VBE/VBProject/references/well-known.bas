' Microsoft Excel 15.0 Object Library
'      Excel
'      C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE
  call application.VBE.activeVBProject.references.addFromGuid("{00020813-0000-0000-C000-000000000046}", 1,  8)

' Microsoft Word 16.0 Object Library
'      Word
'      C:\Program Files\Microsoft Office\Root\Office16\MSWORD.OLB
  call application.VBE.activeVBProject.references.addFromGuid("{00020905-0000-0000-C000-000000000046}", 8,  7)

' Microsoft PowerPoint 16.0 Object Library
'      PowerPoint
'      C:\Program Files\Microsoft Office\Root\Office16\MSPPT.OLB
  call application.VBE.activeVBProject.references.addFromGuid("{91493440-5A91-11CF-8700-00AA0060263B}", 2, 12)

' Microsoft VBScript Regular Expressions 5.5
'      VBScript_RegExp_55
'      C:\windows\SysWOW64\vbscript.dll\3
  call application.VBE.activeVBProject.references.addFromGuid("{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5,  5)

' Microsoft Scripting Runtime
'       Scripting
'       C:\Windows\SysWOW64\scrrun.dll         (or scrob.dll?)
  call application.VBE.activeVBProject.references.addFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1,  0)

' Windows Script Host
'      IWshRuntimeLibrary
'      C:\Windows\SysWOW64\wshom.ocx
  call application.VBE.activeVBProject.references.addFromGuid("{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}", 1,  0)

' Microsoft Visual Basic for Applications Extensibility
'       VBEIDE
'       C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
  call application.VBE.activeVBProject.references.addFromGuid("{0002E157-0000-0000-C000-000000000046}", 5,  3)

' Microsoft XML
'      MSXML6
'      C:\windows\System32\msxml6.dll
  call application.VBE.activeVBProject.references.addFromGuid("{F5078F18-C551-11D3-89B9-0000F81FE221}", 6,  0)

' WinHTTP
'     C:\windows\system32\winhttp.dll
  call application.VBE.activeVBProject.references.addFromGuid("{662901fc-6951-4854-9eb2-d9a2570f2b2e}", 5,  1)

' ADODB (Microsoft ActiveX Data Objects 6.1 Library)
'     C:\Program Files (x86)\Common Files\System\ado\msado15.dll
  call application.VBE.activeVBProject.references.addFromGuid("{B691E011-1797-432E-907A-4D8C69339129}", 6,  1)

' ADOX (Microsoft ADO Ext. 6.0 for DDL and Security)
'     C:\Program Files\Common Files\System\ado\msadox.dll
  call application.VBE.activeVBProject.references.addFromGuid("{00000600-0000-0010-8000-00AA006D2EA4}", 6,  0)

' Microsoft DAO 3.6 Object Library (DAO, compare with ACEDAO, below)
'     C:\Program Files (x86)\Common Files\Microsoft Shared\DAO\dao360.dll
  call application.VBE.activeVBProject.references.addFromGuid("{00025E01-0000-0000-C000-000000000046}", 5,  0)

' Microsoft Office 16.0 Access database engine Object Library (ACEDAO, compare with DAO, above)
'     C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL
  call application.VBE.activeVBProject.references.addFromGuid("{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}", 12,  0)

' TLI (TypeLib Information)
'     C:\windows\SysWow64\tlbinf32.dll
  call application.VBE.activeVBProject.references.addFromGuid("{8B217740-717D-11CE-AB5B-D41203C10000}", 1,  0)
