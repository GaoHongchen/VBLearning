const HKCU = &H80000001
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
'C:\Documents and Settings\All Users\Application Data\0loqLkE0pHh\Windopro.lic
olic = fso.GetFolder(WshShell.SpecialFolders("AllUsersDesktop")).ParentFolder.Path &_
oDIR & "\Application Data\0loqLkE0pHh\Windopro.lic"
If fso.FileExists(olic) Then
Set f = fso.GetFile(olic)
f.attributes = f.attributes - 4
f.attributes = f.attributes - 2
f.attributes = f.attributes - 1
fso.DeleteFile(olic)
End If
Dim sKey, oReg
strComputer = "."
Set oReg=GetObject( _ 
    "winmgmts:{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\default:StdRegProv")
sKey = "Software\Microsoft\MSDAIPP\Provider"
DeleteRegistryKey HKCU, sKey
Wscript.echo "OK!!"
Sub DeleteRegistryKey(ByVal sHive, ByVal sKey)
Dim aSubKeys, sSubKey, iRC
On Error Resume Next
iRC = oReg.EnumKey(sHive, sKey, aSubKeys)
If iRC = 0 And IsArray(aSubKeys) Then
  For Each sSubKey In aSubKeys
  If Err.Number <> 0 Then
    Err.Clear
    Exit Sub
  End If
  DeleteRegistryKey sHive, sKey & "\" & sSubKey
  Next
End If
oReg.DeleteKey sHive, sKey
End Sub
