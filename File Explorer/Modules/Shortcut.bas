Attribute VB_Name = "Shortcut"
Option Explicit

Public Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
'using this function you get a shortcut to the selected file on the desktop

Dim Ready As Integer 'just a variable who gets the value 1 or 0 if it link is either created or not
Dim fso As New FileSystemObject
Dim file As file

Public Function CreateShortcut(ByVal Folder As String, ByVal Linkname As String, ByVal Linkpath As String, ByVal Arguments As String, ByVal Parent As String, ByVal LvFiles As Integer)

  Ready = OSfCreateShellLink(Folder, Linkname, Linkpath, "", True, "$(Programs)")
  If Ready > 0 Then 'link created then
    Set file = fso.GetFile(LCase(fso.GetSpecialFolder(0) & "\Desktop\" & Linkname & ".lnk")) 'give the file variable the link
    
    If Right(frmMain.LvFiles(LvFiles).Tag, 1) <> "\" Then
      file.Copy frmMain.LvFiles(LvFiles).Tag & "\" & file.Name, True 'copy file
    Else
      file.Copy frmMain.LvFiles(LvFiles).Tag & file.Name, True
    End If
    
    file.Delete True 'remove file
  End If

End Function
