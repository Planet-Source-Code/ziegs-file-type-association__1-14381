<div align="center">

## File Type Association


</div>

### Description

Create file extension for your application.
 
### More Info
 
The name and extension of your app

You need a Rich Text Box to use this!!

The association

A new file extension. (Good side effect)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ziegs](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ziegs.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ziegs-file-type-association__1-14381/archive/master.zip)





### Source Code

```
'//							File Association
'//I made this to figure out how associate a file extension with a project I am currently '//working on called ZWord. I wanted '//the .zwd extension, so this is what I did.
'//Goes Under General Declarations for Main Form
'// Registry windows api calls
Private Declare Function RegCreateKey& Lib "advapi32.DLL" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, lphKey As Long)
Private Declare Function RegSetValue& Lib "advapi32.DLL" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpszSubKey As String, ByVal fdwType As Long, ByVal lpszValue As String, ByVal dwLength As Long)
'// Required constants
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 256&
Private Const REG_SZ = 1
'// procedure you call to associate the zwd extension with your program.
Private Sub MakeDefault()
  Dim sKeyName As String '// Holds Key Name in registry.
  Dim sKeyValue As String '// Holds Key Value in registry.
  Dim ret    As Long  '// Holds error status if any from API calls.
  Dim lphKey  As Long  '// Holds created key handle from RegCreateKey.
  '// This creates a Root entry called "ZWord"
  sKeyName = "ZWord" '// Application Name
  sKeyValue = "Zword Document" '// File Description
  ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
  ret = RegSetValue&(lphKey&, Empty, REG_SZ, sKeyValue, 0&)
  '// This creates a Root entry called .zwd associated with "ZWord".
  sKeyName = ".zwd" '// File Extension
  sKeyValue = "ZWord" '// Application Name
  ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
  ret = RegSetValue&(lphKey, Empty, REG_SZ, sKeyValue, 0&)
  '//This sets the command line for "ZWord".
  sKeyName = "Zword" '// Application Name
  If App.Path Like "*\" Then
    sKeyValue = App.Path & App.EXEName & ".exe %1" '// Application Path
  Else
    sKeyValue = App.Path & "\" & App.EXEName & ".exe %1" '// Application Path
  End If
  ret = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey)
  ret = RegSetValue&(lphKey, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
End Sub
'//Stick This into the Form or MDIForm Load
  '// ensure we only register once. When debugging etc, remove the SaveSetting line, so your program will
  '// always attempt to register the file extension.
  If GetSetting(App.Title, "Settings", "RegisteredFile", 0) = 0 Then
    '// associate tmg extension with this app
    MakeDefault
    SaveSetting App.Title, "Settings", "RegisteredFile", 1
  End If
'// If you are in an MDI App, then put this in
'// MDIForm_Load:
If Command = "" Then
  Resume Next
Else
  frmMain.ActiveForm.rtfText.LoadFile Command
End If
'// If you are in a SDI App, put this in Form_Load
If Command = "" Then
  Resume Next
Else
  frmMain.rtfText.LoadFile Command
End If
```

