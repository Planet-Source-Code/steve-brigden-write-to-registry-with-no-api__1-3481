<div align="center">

## Write to Registry with No API


</div>

### Description

Who says you cant write to the registry without API?

This code is fast,smooth, and refreshes the registry instantly.

This is speed at work!
 
### More Info
 
just create a 'Reg File' call it Reg.Reg' for simplicity in keeping with the code.

Place it in the root dir: eg: c:\reg.Reg

paste the code straight into your form and away you go.

This function is great!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[steve brigden](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-brigden.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-brigden-write-to-registry-with-no-api__1-3481/archive/master.zip)

### API Declarations

Not Likely


### Source Code

```
Function SetDwordKeyValue(Hkey As String, SubKey As String, Keyname As String, Dword As String, Value As String)
Dword = "=dword:"
A$ = "REGEDIT4" & vbCrLf & "[" & Hkey & "\" & SubKey & "]" & vbCrLf & """" & Keyname & """" & Dword & Value
Open "c:\reg.reg" For Output As 1'create a 'Reg file and name it: 'Reg.reg
Print #1, A$
Close #1
ret = Shell("regedit.exe /s " & "c:\reg.reg", 0)
Kill "c:\reg.reg"
End Function
Sub DoIt(rtn As Boolean)
'Disable/Re-enable Regedit.exe
If rtn = True Then
ret = SetDwordKeyValue("HKEY_CURRENT_USER", "\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", "", "00000001")
Else
ret = SetDwordKeyValue("HKEY_CURRENT_USER", "\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", "", "00000000")
End If
End Sub
Private Sub Form_Load()
'Writing to the Registry with no API's! What! No joke, It can be done!
'Changing this value to True Disables Regedit.Exe & Vice Versa.
'Comments:Steve.Brigden@Usa.Net
DoIt False
End Sub
```

