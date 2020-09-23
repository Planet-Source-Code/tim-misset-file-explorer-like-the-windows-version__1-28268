Attribute VB_Name = "Properties"
Option Explicit
Public Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
'api function, used to execute the info

Public Type SHELLEXECUTEINFO 'Api Type, contains all file information
        cbSize As Long
        fMask As Long
        Hwnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type

'just constants again
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40


Public Sub GetPropertiesPopup(Path As String)

  Dim SEI As SHELLEXECUTEINFO
  With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .Hwnd = frmMain.Hwnd
    .lpVerb = "properties" 'commando so the right pop shows up
    .lpFile = Path
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
  End With
    
  Call ShellExecuteEX(SEI) 'executes the information
  
End Sub
