Option Explicit

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Early Binding ensures that the object and its methods are verified at compile time
If objFSO.FolderExists("C:\MyFolder") Then
  objFSO.DeleteFolder "C:\MyFolder", True
  WScript.Echo "Folder deleted successfully."
Else
  WScript.Echo "Folder does not exist."
End If

Set objFSO = Nothing