<div align="center">

## CreateFolders


</div>

### Description

Make nested subfolders in a single method.
 
### More Info
 
sPath: Fully-qualified absolute or relative path you wish to create.

'Example 1: \\NetworkVolume\NetworkShare\ExistingDir\NewDir\NewSubdir\NewSubDir\

'Example 2: C:\Program Files\a\b\c\d\e\f\g\h\i\j\k\l\m\n\o\p\q\r\s\t\u\v\w\x\y\z

Add reference "Microsoft Scripting Runtime" (scrrun.dll) available at http://www.microsoft.com/scripting or with VB6.

Number 70, Permission Denied error will occur is write access or directory create access is not allowed for the drive for that user.

'Reference "Microsoft Scripting Runtime" (scrrun.dll) available at http://www.microsoft.com/scripting or with VB6.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gerald Bryant](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gerald-bryant.md)
**Level**          |Unknown
**User Rating**    |4.2 (162 globes from 39 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gerald-bryant-createfolders__1-2179/archive/master.zip)

### API Declarations

Add reference "Microsoft Scripting Runtime" (scrrun.dll) available at http://www.microsoft.com/scripting or with VB6.


### Source Code

```
Public Sub CreateFolders(ByVal sPath As String)
 Dim oFileSystem As New Scripting.FileSystemObject
 'or late-bind with:
 'Dim oFileSystem As Object
 'Set oFileSystem = CreateObject("Scripting.FileSystemObject")
 On Error GoTo ErrorHandler
 With oFileSystem
  ' Is this drive valid and ready?
  If .DriveExists(.GetDriveName(sPath)) Then
   ' Is this folder not yet valid?
   If Not .FolderExists(sPath) Then
    ' Recurse back in to this method until a parent folder is valid.
    CreateFolders .GetParentFolderName(sPath)
    ' Create only a nonexistant folder before exiting the method.
     .CreateFolder sPath
   End If
  End If
 End With
 Set oFileSystem = Nothing
ExitMethod:
 Exit Sub
ErrorHandler:
 App.LogEvent "CreateFolders Error in " & Err.Source & _
 ": Could not create " & sPath & ".", vbLogEventTypeInformation
End Sub
```

