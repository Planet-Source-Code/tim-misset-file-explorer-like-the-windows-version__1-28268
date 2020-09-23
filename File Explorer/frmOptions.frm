VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Select View Right List"
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   6855
      Begin VB.OptionButton optView2 
         Caption         =   "Large Icons"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optView2 
         Caption         =   "Small Icons"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optView2 
         Caption         =   "List"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optView2 
         Caption         =   "Details"
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   13
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtExcludefiles 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.CheckBox chkExcludeBut 
      Caption         =   "Exclude all files from list execept if extension is"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   5175
   End
   Begin VB.CheckBox chkExclude 
      Caption         =   "Exclude files with the following extension from lists"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Exclude Files"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   6855
   End
   Begin VB.OptionButton optView 
      Caption         =   "Details"
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   4
      Top             =   720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optView 
      Caption         =   "List"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton optView 
      Caption         =   "Small Icons"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton optView 
      Caption         =   "Large Icons"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame frmViewLeft 
      Caption         =   "Select View Left List"
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()

  Dim optbutton As OptionButton
  For Each optbutton In optView
    If optbutton.Value = True Then
      frmMain.LvFiles(0).View = optbutton.Index
    End If
  Next
  For Each optbutton In optView2
    If optbutton.Value = True Then
      frmMain.LvFiles(1).View = optbutton.Index
    End If
  Next
  'Look at wich optionbutton is enabled.
  'if it is found, the index matched the value of the lvfiles(index).view
  'i.e. Iconview = 0 so the Icon optionbutton's index = 0
  'SmallIconview = 1
  'List = 2
  'Details = 3
    
  If chkExclude.Value = 1 Then
    frmMain.Exclude = txtExcludefiles
  Else
    frmMain.Exclude = ""
  End If
  If chkExcludeBut.Value = 1 Then
    frmMain.Excludebut = txtExcludefiles
  Else
    frmMain.Excludebut = ""
  End If
  'Set the exclude variables on the main form
  'so it filters files on loading
  
  With frmMain
    .Load_lvFiles 0, .LvFiles(0).Tag, .Exclude, .Excludebut
    .Load_lvFiles 1, .LvFiles(1).Tag, .Exclude, .Excludebut
  End With
  'refresh the lists on the main form both
  
End Sub

Private Sub cmdCancel_Click()

  optView(frmMain.LvFiles(0).View).Value = True
  optView2(frmMain.LvFiles(1).View).Value = True
  'Turn the optiobuttons back to the original state
  'by watching the current view on the lvfile.view
  
  If frmMain.Exclude = "" Then
    chkExclude.Value = 0
  Else
    chkExclude.Value = 1
    txtExcludefiles.Text = frmMain.Exclude
  End If
  If frmMain.Excludebut = "" Then
    chkExcludeBut.Value = 0
  Else
    chkExcludeBut.Value = 1
     txtExcludefiles.Text = frmMain.Excludebut
  End If
  'see if there is an exclude or excludebut currently active
  'if so then the setting will be changed to that exclude
  
  If chkExclude.Value = 0 And chkExcludeBut.Value = 0 Then
    txtExcludefiles.Enabled = False
    txtExcludefiles.BackColor = &H80000004
  End If
  'if no exclude checkbox is checked then disable the textbox
  
  Me.Hide
  'DON'T close the form
  'by putting it in the hide mode, it will maintain the settings
  'if it is closed then the standard settings will show everytime
  'the form is shown

End Sub

Private Sub cmdOk_Click()

  cmdApply_Click
  Me.Hide
  'apply changes and hide form

End Sub

Private Sub chkExclude_Click()

  If chkExclude.Value = 1 Then
    chkExcludeBut.Value = 0 'make sure only one of the two or none is clicked.
    txtExcludefiles.Enabled = True 'enable cause one of the two checkboxes is checked
    txtExcludefiles.BackColor = &H80000005 'change color to white
  End If
  If chkExclude.Value = 0 And chkExcludeBut.Value = 0 Then
    txtExcludefiles.Enabled = False 'if none checked then disable the textbox
    txtExcludefiles.BackColor = &H80000004 'gray the box
  End If
    
End Sub

Private Sub chkExcludeBut_Click()

  If chkExcludeBut.Value = 1 Then
    chkExclude.Value = 0
    txtExcludefiles.Enabled = True
    txtExcludefiles.BackColor = &H80000005
  End If
  If chkExclude.Value = 0 And chkExcludeBut.Value = 0 Then
    txtExcludefiles.Enabled = False
    txtExcludefiles.BackColor = &H80000004
  End If
  'same as the chkExclude code
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Cancel = 1 'cancel the unload procedure
  cmdCancel_Click 'use the cancel procedure instead so the form will be hidden, not closed
  
End Sub
