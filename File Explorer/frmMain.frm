VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   ".: File Explorer :. ©2001 by Tim Misset"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CDBox 
      Left            =   960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSeperator 
      Height          =   6615
      Left            =   5040
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6615
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   1200
      Width           =   15
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   7830
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   741
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicFiles32 
      BackColor       =   &H80000009&
      Height          =   600
      Left            =   720
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox PicFiles16 
      BackColor       =   &H80000009&
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComctlLib.ImageList ILFiles32 
      Index           =   0
      Left            =   720
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ILFiles16 
      Index           =   0
      Left            =   120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LvFiles 
      Height          =   6615
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11668
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "frmMain.frx":0000
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ext"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date & Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ILDrives 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":114E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbDrives 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ILDrives"
      DisabledImageList=   "ILQuickLaunch"
      _Version        =   393216
   End
   Begin VB.PictureBox PicQuickLaunch 
      Height          =   300
      Left            =   600
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComctlLib.ImageList ILQuickLaunch 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbQuickLaunch 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      DisabledImageList=   "ILDrives"
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LvFiles 
      Height          =   6615
      Index           =   1
      Left            =   5040
      TabIndex        =   6
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   11668
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ext"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date & Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ILFiles32 
      Index           =   1
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ILFiles16 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "N&ew"
         Begin VB.Menu mnuFileNewShortcut 
            Caption         =   "Short&cut"
         End
         Begin VB.Menu mnuFileNewFolder 
            Caption         =   "&Folder"
         End
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCreateShortcut 
         Caption         =   "Create Short&cut"
      End
      Begin VB.Menu mnuFileRemove 
         Caption         =   "Remo&ve"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Re&name"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuEdit_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuEditTurnselection 
         Caption         =   "&Turn Selection"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$  This code is written by Tim Misset   $
'$  It is used for learning only         $
'$  any questions: timmisset@hotmail.com $
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Icons
Dim hSIcon As Long 'SmallIcon
Dim hLIcon As Long 'LargeIcon
Dim imgObj As ListImage 'Image in ImageList
Dim r As Long 'just a common variable so an action can happen
Public Exclude As String
Public Excludebut As String
'FSO
Dim fso As New FileSystemObject 'Contains info for files, maps, drives and textstream
Dim file As file 'a file variable
Dim Folder As Folder 'a folder variable
Dim Drive As Drive 'a drive variable
'More
Dim Button As Button 'a button on the toolbar
Dim Item As ListItem 'an item in de filelist
Dim SelectedList As Integer 'needed to know which list we deal with
Dim oldSortkey As Integer 'in the Load_lvFilesIcon we change the sortkey, this variable is needed to return the sortkey
Dim oldSortorder As Integer 'same as sortkey
Dim lvFilesWidth As Single 'maintains what percentage of the screen is reserved for the first list


Private Sub Form_Load()
  lvFilesWidth = 0.5 'on start both lists will have the same size
  Load_tbQuickLaunch 'call a procedure
  Load_tbDrives 'idem
  Load_lvFiles 0, CurDir("c"), Exclude, Excludebut 'idem
  Load_lvFiles 1, CurDir("c"), Exclude, Excludebut 'idem
  
  SelectedList = 0
  LvFiles_ItemClick SelectedList, LvFiles(SelectedList).SelectedItem
  
End Sub

Private Sub Load_tbQuickLaunch()
On Error Resume Next
Dim strQuickLaunchPath As String 'make a stringvariable
strQuickLaunchPath = fso.GetSpecialFolder(WindowsFolder) & "\Application Data\Microsoft\Internet Explorer\Quick Launch"
'strQuickLaunchPath now contains the path to the Windows Quicklaunch Dir

If fso.FolderExists(strQuickLaunchPath) Then 'make sure the path is correct, to prevent an error
  Set Folder = fso.GetFolder(strQuickLaunchPath)
  For Each file In Folder.Files 'call the same procedure for each file in the folder
    Set Button = tbQuickLaunch.Buttons.Add 'sets the button variable as a new button on the quicklaunch toolbar
    Button.ToolTipText = fso.GetBaseName(file.Path) 'tooltiptext = message you see when you hold the mousebutton over the button (now contains the basename = filename - extensionname)
    Button.Tag = file.Path 'tag = extra information for an object, now it contains the path to the file
    
    hSIcon = SHGetFileInfo(file.Path, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    'hSicon now contains the info what icon is connected to file
    'it uses the SHGetFileInfo Function to do so... (the function is in the icons module)
    If hSIcon <> 0 Then 'make sure the hSicon contains info (0 means an nothing, anything but 0 is good)
      With PicQuickLaunch
        Set .Picture = LoadPicture("") 'the picturebox will now be ready to load an image
        .AutoRedraw = True 'AutoRedraw so it fits better
        r = ImageList_Draw(hSIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
        'Calls the ImageList_draw function to draw the icon into the picturebox
        .Refresh 'refresh the picturebox with the new image
      End With
      Set imgObj = ILQuickLaunch.ListImages.Add(Button.Index, , PicQuickLaunch.Image)
      'imgobj will be the new image in the list. it is the image from the picturebox
    End If
  Next
  tbQuickLaunch.ImageList = ILQuickLaunch 'after all files are handled you link the imagelist to the toolbar
  
  For Each Button In tbQuickLaunch.Buttons
    Button.Image = Button.Index 'since we gave the images in the imagelist the same indexnumber as the buttons
    'the button.image (integer) will be the same as the buttons index
  Next
End If

End Sub

Private Sub Form_Resize()
  With LvFiles
    .Item(0).Top = tbDrives.Top + tbDrives.Height
    .Item(0).Left = 0
    .Item(0).Height = Me.Height - sbMain.Height - tbDrives.Height - tbQuickLaunch.Height - 700
    .Item(0).Width = Me.Width * lvFilesWidth
    
    'the above code is used to make sure the first list will fit.
    'for the width is uses the lvfileswidth
    '(on form_load this will be 0.5 so it will be half of the screen)
    'if the seperator between the lists is moved it makes sure the
    'ratio will stay the same
    
    picSeperator.Left = .Item(0).Width
    picSeperator.Height = .Item(0).Height
    picSeperator.Top = .Item(0).Top
    'equalizes the Seperator to the first list, except for the width offcourse
    
    .Item(1).Top = .Item(0).Top
    .Item(1).Left = .Item(0).Width + picSeperator.Width
    .Item(1).Width = Me.Width - .Item(0).Width - picSeperator.Width - 100
    .Item(1).Height = .Item(0).Height
    'equalizes the second list to the first list, except for the width
    'the width is the remaining space available
  End With
End Sub


Private Sub LvFiles_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)
On Error GoTo Errhandler:
  Set Item = LvFiles(Index).SelectedItem
  If Item.Text = "New Folder" Then 'what if we handle a new folder
    Item.Text = NewString 'adjust the itemtext to the newtext
    
    If Right(LvFiles(Index).Tag, 1) <> "\" Then
      fso.CreateFolder (LvFiles(Index).Tag & "\" & NewString)
    Else
      fso.CreateFolder (LvFiles(Index).Tag & NewString)
    End If
    'this above code makes the new list and places the "\" if needed
  Else
    If Item.Text = "[..]" Then
      Exit Sub
    End If
    'the "[..]" can't be renamed so we need to prevent any attempts
    
    If fso.FolderExists(Item.Text) Then
      Set Folder = fso.GetFolder(Item.Text)
      Folder.Name = NewString
      'if the selecteditem is a folder, change the foldername
    ElseIf fso.FileExists(Item.Text) Then
      Set file = fso.GetFile(Item.Text)
      file.Name = NewString
      'if the selecteditem is file, change the filename
    End If
  End If
  
  Load_lvFiles Index, LvFiles(Index).Tag, Exclude, Excludebut
  'when ready then refresh the list
Exit Sub

Errhandler:
Load_lvFiles Index, LvFiles(Index).Tag, Exclude, Excludebut

'if an error occurs (like foldername already exists) then
'then just refresh the list and no one will see any error message

End Sub

Private Sub LvFiles_BeforeLabelEdit(Index As Integer, Cancel As Integer)

  Set Item = LvFiles(Index).SelectedItem
    If Item.Text = "[..]" Then
      Cancel = 1
    End If
    
  'this code makes sure you can't edit the text of the "[..]" item
  'this is needed because it's a fake folder, it doesn't exist and is only a link
  'for exploring one level up.
  'if the name would be edited a crash be a sure thing

End Sub

Private Sub LvFiles_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  
  If LvFiles(Index).SortOrder = 0 Then
    Load_SortList Index, ColumnHeader.Index - 1, 1
    'calls the Load_Sortlist with the right info
    'if the sortorder is 0 the new sortorder is 1
  Else
    Load_SortList Index, ColumnHeader.Index - 1, 0
    'idem
  End If
End Sub

Private Sub LvFiles_DblClick(Index As Integer)

  Set Item = LvFiles(Index).SelectedItem 'sets the item as the item selected in the selected list
  If fso.FileExists(Item.Tag) Then
    ShellExecute Me.Hwnd, "open", Item.Tag, "", "", 3
    'if the selected item represents a file then the file is executed
    'calling the Api-function ShellExcute
  Else
    Load_lvFiles Index, Item.Tag, Exclude, Excludebut
    'if it isn't a file then it must be a folder
    'so the list will go the dir
  End If

End Sub

Private Sub LvFiles_GotFocus(Index As Integer)
  SelectedList = Index 'when one of the two lists is selected then
  'the selectedlist variable will know which one
  'the selectedlist is used for functions who can't detect themself which list is selected
  
  sbMain.SimpleText = LvFiles(Index).Tag
End Sub

Private Sub LvFiles_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)

  If fso.FileExists(Item.Tag) = True Then
    Set file = fso.GetFile(Item.Tag)
    mnuFileOpen.Caption = "Open " & file.Name
  ElseIf fso.FolderExists(Item.Tag) = True Then
    Set Folder = fso.GetFolder(Item.Tag)
    mnuFileOpen.Caption = "Explore " & Folder.Name
  End If
  'no comment needed, I guess
End Sub

Private Sub LvFiles_OLEDragDrop(Index As Integer, Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim Index2 As Integer

  If Index = 1 Then
    Index2 = 0
  Else
    Index2 = 1
  End If
  
  'first make two different index numbers
  'index = index for destinationlist
  'index2 = index for sourcelist
  
  Dim target As String
  If Right(LvFiles(Index).Tag, 1) <> "\" Then
    target = LvFiles(Index).Tag & "\"
  Else
    target = LvFiles(Index).Tag
  End If
  'determine the targetdirectory
  
  For Each Item In LvFiles(Index2).ListItems
    If Item.Selected = True Then
      If fso.FileExists(Item.Tag) Then
        fso.CopyFile Item.Tag, target & file.Name, True
        'copies the file to the current folder of the other list
      ElseIf fso.FolderExists(Item.Key) Then
        If Item.Text <> "[..]" Then 'make sure the fake folder won't be copied
          fso.CopyFolder Item.Tag, target & Folder.Name, True
          'copies the folder to the current folder of the other list
        End If
      End If
    End If
  Next

  'after the files/folders are copied
  'refresh the list again and set the mouseicons to custom (=0)
  
  Load_lvFiles Index, LvFiles(Index).Tag, Exclude, Excludebut
  LvFiles(0).MousePointer = 0
  LvFiles(1).MousePointer = 0
  

End Sub

Private Sub mnuEditOptions_Click()

  frmOptions.Show

End Sub

Private Sub mnuEditSelectAll_Click()

  For Each Item In LvFiles(SelectedList).ListItems
    Item.Selected = True
  Next

End Sub

Private Sub mnuEditUnselectAll_Click()

  

End Sub

Private Sub mnuEditTurnselection_Click()

  For Each Item In LvFiles(SelectedList).ListItems
    If Item.Selected = True Then
      Item.Selected = False
    Else
      Item.Selected = True
    End If
  Next

End Sub

Private Sub mnuFileCreateShortcut_Click()

  Set Item = LvFiles(SelectedList).SelectedItem
    If fso.FileExists(Item.Tag) Then
      CreateShortcut "..\..\Desktop", file.Name, file.Path, "", file.ParentFolder, SelectedList
    ElseIf fso.FolderExists(Item.Tag) Then
      CreateShortcut "..\..\Desktop", Folder.Name, Folder.Path, "", Folder.ParentFolder, SelectedList
    End If
  
  frmMain.Load_lvFiles SelectedList, LvFiles(SelectedList).Tag, Exclude, Excludebut  'refresh the screen
  'Look at the create newshortcut for comments

End Sub

Private Sub mnuFileExit_Click()
  Unload Me
  Unload frmOptions
  'unloads the form me...
  'me is the current form
End Sub

Private Sub mnuFileNewFolder_Click()

  Set Item = LvFiles(SelectedList).ListItems.Add
    Item.Text = "New Folder" 'standard name for new folder
    Item.ListSubItems.Add 1, , " "
    Item.ListSubItems.Add 2, , " "
    Item.ListSubItems.Add 3, , Now 'current system time
    
  LvFiles(SelectedList).Refresh 'refresh list so the item is in it permantly (unless the list is cleared)
  
  Item.Selected = True 'select new item
  LvFiles(SelectedList).StartLabelEdit 'start the labeledit
  
End Sub

Private Sub mnuFileNewShortcut_Click()

On Error GoTo Errhandler

With CDBox 'CDBox = Common Dialog Box
  .DialogTitle = "Select a file to create a link to" 'Title of the box
  .CancelError = True 'if cancelled then give an error
  .ShowOpen 'open the dialog box
End With
Set file = fso.GetFile(CDBox.FileName) 'set the file variable to the file returned from the box
Set Folder = fso.GetFolder(LvFiles(SelectedList).Tag) 'set the folder as the active folder of the selected list

CreateShortcut "..\..\Desktop", file.Name, file.Path, "", Folder.Path, SelectedList
'run the createshortcut function
Exit Sub
  
Errhandler:
  Exit Sub

End Sub

Private Sub mnuFileOpen_Click()

  LvFiles_DblClick SelectedList
  'the same code is needed as in the dblclick so just call that code

End Sub

Private Sub mnuFileProperties_Click()

  Set Item = LvFiles(SelectedList).SelectedItem
  GetPropertiesPopup Item.Tag
  'use the propertiesmodule to get the propertiespopup

End Sub

Private Sub mnuFileRemove_Click()

On Error GoTo Errhandler

Dim intSelected As Integer
intSelected = 0
'this variable is used to count the amount of files selected

  For Each Item In LvFiles(SelectedList).ListItems
    If Item.Selected = True Then
      If Item.Text <> "[..]" Then 'dont count the fake folder ("[..]")
        intSelected = intSelected + 1
      End If
    End If
  Next
  
  If MsgBox("Are you sure you want to delete " & intSelected & " item(s)?", vbYesNo + vbQuestion, Caption) = vbNo Then
    'if the users changes his mind on deleting the files/folders
    'exit the sub so nothing happends
    Exit Sub
  End If
  
  For Each Item In LvFiles(SelectedList).ListItems
    'check each file again and if selected
    'make sure it isn't the fake folder ("[..]")
    If Item.Selected = True Then
      If Item.Text = "[..]" Then
        'if the item is the fake folder, then just go to the next
        Resume Next
      End If
      
      If fso.FileExists(Item.Tag) Then 'if the item represents a file, delete the file
        fso.DeleteFile Item.Tag, True
      ElseIf fso.FolderExists(Item.Tag) Then 'idem with folder
        fso.DeleteFolder Item.Tag, True
      End If
    End If
  Next
  
  'after deleting refresh the current folder
  Load_lvFiles SelectedList, LvFiles(SelectedList).Tag, Exclude, Excludebut
  Exit Sub
  
Errhandler: 'on error show the error
  MsgBox Err.Number & Space(10) & Err.Description

End Sub

Private Sub mnuFileRename_Click()

  LvFiles(SelectedList).StartLabelEdit
  'no comment

End Sub

Private Sub picSeperator_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then 'make sure the first mousebutton is presseddown
    If picSeperator.Left + x > 1000 And picSeperator.Left + x < (Me.Width - 1000) Then
      'make sure the list won't be smaller then 1000 units on both sides
      
      picSeperator.Left = picSeperator.Left + x 'moves the picseperator along with the mouse
      LvFiles(0).Width = picSeperator.Left 'changes the width of the list along with the seperator
      LvFiles(1).Left = picSeperator.Left + picSeperator.Width 'changes the position of the second list along with the seperator
      LvFiles(1).Width = Me.Width - LvFiles(0).Width - picSeperator.Width - 100 'changes the width of the list along with the seperator
      lvFilesWidth = LvFiles(0).Width / Me.Width
      'puts a new value to the lvfileswidth (used by the form_resize)
      
    End If
  End If
End Sub

Private Sub tbDrives_ButtonClick(ByVal Button As MSComctlLib.Button)
  Set Drive = fso.GetDrive(Button.Tag)
  If Drive.IsReady = True Then 'make sure the station contains info
    Load_lvFiles SelectedList, Button.Tag, Exclude, Excludebut
  Else
    If MsgBox("Can't load station!", vbRetryCancel + vbExclamation, Caption) = vbRetry Then
      tbDrives_ButtonClick Button
    End If
  End If
  'Loads the CURRENT drive path in the selected list
  'the differents between the currentpath and the root is
  'root: "c:\" current: "c:" then it searches for the active folder
  
  'if however, the drive isn't ready it shows a messagebox containing a
  'retry and cancel button.
  'if retry is pressed, we will........ retry, duh!
  
End Sub

Private Sub tbQuickLaunch_ButtonClick(ByVal Button As MSComctlLib.Button)
  ShellExecute Me.Hwnd, "open", Button.Tag, "", "", 3
  'same as the filehandle in the lists
  'it uses the shellexecute to open the right program with the link
End Sub

Private Sub Load_tbDrives()
  For Each Drive In fso.Drives
    Set Button = tbDrives.Buttons.Add
    'sets the button variable as a new button
    Button.Caption = LCase(Drive.DriveLetter)
    'sets the button.caption as a lowcase letter of the drive
    If Drive.IsReady Then 'IsReady = contains data
      Button.ToolTipText = Drive.VolumeName
      'sets the tooltext to the volumename if the drive is ready
    Else
      Button.ToolTipText = LCase(Drive.DriveLetter) & ": - No data inserted"
      'if no data then there's no volumename
      'so then it will contain de the driveletter and some extra comment
    End If
    Button.Image = Drive.DriveType
    'the drivetype is an integer and in the imagelist
    'there are some icons for drivetypes already sorted so
    'the equal with the drivetype variable
    Button.Tag = Drive.DriveLetter & ":"
    'tag = extra info, it contains the path to the currentpath on the drive
    'to make sure you always go to the root, change ":" in ":\"
  Next
End Sub

Public Sub Load_lvFiles(Index As Integer, Path As String, Exclude As String, Excludebut As String)

LvFiles(Index).ListItems.Clear
LvFiles(Index).SmallIcons = Nothing
LvFiles(Index).Icons = Nothing
ILFiles16(Index).ListImages.Clear
ILFiles32(Index).ListImages.Clear
'the above code makes sure all lists are clear and unlinked to all objects
  
  ChDir Path 'change the current path of a drive
  
  LvFiles(Index).Tag = CurDir(Path) 'get the real path (not c: but c:\program files\myprogram, for example)
  'add the active path to the lvfiles extra info
  'so we can always see it's location
  
  sbMain.SimpleText = LvFiles(Index).Tag
    
Dim subFolder As Folder 'make a second folder variable
  Set Folder = fso.GetFolder(Path) 'set the variable so it contains the info for the path we want to explore
  
  For Each subFolder In Folder.SubFolders
    Set Item = LvFiles(Index).ListItems.Add
    Item.Text = "[" & subFolder.Name & "]"
    Item.Bold = True
    Item.ListSubItems.Add , , " "
    Item.ListSubItems.Add 2, , " "
    Item.ListSubItems.Add(3).Text = FileDateTime(subFolder.Path)
    Item.Tag = subFolder.Path
  Next
  'the above code add's all subfolder to the list
  
  For Each file In Folder.Files
    If Exclude <> "" Then 'if something must be excluded, exclude it
      If LCase(fso.GetExtensionName(file.Path)) <> LCase(Exclude) Then
        Set Item = LvFiles(Index).ListItems.Add
        Item.Text = fso.GetBaseName(file.Path)
        Item.Bold = False
        Item.ListSubItems.Add 1, , fso.GetExtensionName(file.Path)
        Item.ListSubItems.Add 2, , Round(file.Size / 1024, 0) & " kB"
        Item.ListSubItems.Add 3, , FileDateTime(file.Path)
        Item.Tag = file.Path
      End If
    ElseIf Excludebut <> "" Then 'same as exclude
      If LCase(fso.GetExtensionName(file.Path)) = LCase(Excludebut) Then
        Set Item = LvFiles(Index).ListItems.Add
        Item.Text = fso.GetBaseName(file.Path)
        Item.Bold = False
        Item.ListSubItems.Add 1, , fso.GetExtensionName(file.Path)
        Item.ListSubItems.Add 2, , Round(file.Size / 1024, 0) & " kB"
        Item.ListSubItems.Add 3, , FileDateTime(file.Path)
        Item.Tag = file.Path
      End If
    Else 'if nothing must be excluded then just load all
      Set Item = LvFiles(Index).ListItems.Add
      Item.Text = fso.GetBaseName(file.Path)
      Item.Bold = False
      Item.ListSubItems.Add 1, , fso.GetExtensionName(file.Path)
      Item.ListSubItems.Add 2, , Round(file.Size / 1024, 0) & " kB"
      Item.ListSubItems.Add 3, , FileDateTime(file.Path)
      Item.Tag = file.Path
    End If
  Next
  'the above code add's all files to the list
  
  
  If Folder.IsRootFolder = False Then
    Set Item = LvFiles(Index).ListItems.Add(1)
    Item.Bold = True
    Item.Text = "[..]"
    Item.Tag = Left(Folder.Path, Len(Folder.Path) - Len(Folder.Name))
    Item.ListSubItems.Add , , " "
    Item.ListSubItems.Add 2, , " "
  End If
  'the above code checks if the folder is the rootfolder
  'if not then it puts an extra "[..]" item to list with index 1 so it shows
  'on top off the list.
  
  Load_lvFilesIcons (Index) 'go to phase 2; adding the icons
  
End Sub

Private Sub Load_lvFilesIcons(Index As Integer)
On Error Resume Next
With LvFiles(Index)
  oldSortkey = .SortKey
  oldSortorder = .SortOrder
  .SortOrder = lvwAscending
  .SortKey = 1
  .Sorted = True
End With
'make sure the lists is sorted on fileextension
'this is needed to make sure we don't get double
'icons in the imagelists (there is a max to the amount of icons to the list)

Dim strExt As String 'Variable that contains the file's extension (like "exe")
Dim strLastExt As String 'Variable that contains the previous file's extension

  For Each Item In LvFiles(Index).ListItems
    If fso.FileExists(Item.Tag) Then 'first we iconize (is that a word?) all files
      Set file = fso.GetFile(Item.Tag)
      strExt = LCase(fso.GetExtensionName(Item.Tag))
      If strExt <> strLastExt Then 'compare the currentfile extension with last
        'if they don't match then add a new icon to the list
        
        hSIcon = SHGetFileInfo(file.Path, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        hLIcon = SHGetFileInfo(file.Path, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        With PicFiles16
          Set .Picture = LoadPicture("")
          .AutoRedraw = True
          r = ImageList_Draw(hSIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
          .Refresh
        End With
        With PicFiles32
          Set .Picture = LoadPicture("")
          .AutoRedraw = True
          r = ImageList_Draw(hLIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
          .Refresh
        End With
        'the above code draw's the icon's in the pictureboxes
        'for details look at the Load_tbQuickLaunch procedure
        
        If strExt <> "exe" And strExt <> "ico" And strExt <> "lnk" Then
          'make sure it isn't an exe or ico file (they must be different each file)
          Set imgObj = ILFiles16(Index).ListImages.Add(, , PicFiles16.Image)
          Set imgObj = ILFiles32(Index).ListImages.Add(, , PicFiles32.Image)
          strLastExt = strExt
        Else
          Set imgObj = ILFiles16(Index).ListImages.Add(, "exe" & Item.Index, PicFiles16.Image)
          Set imgObj = ILFiles32(Index).ListImages.Add(, "exe" & Item.Index, PicFiles32.Image)
          strLastExt = "anything_but_exe_or_ico"
          'the strLastExt should be not a valid extension so every exe and ico and lnk will
          'get it's own private image
        End If
      End If
    ElseIf fso.FolderExists(Item.Tag) = True Then 'now handle the folders
      'all folders should have different icons as well
      'since they don't have extensionname's we just hope there aren't more
      'than 200 subfolders or so in a folder
      hSIcon = SHGetFileInfo(Item.Tag, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
      hLIcon = SHGetFileInfo(Item.Tag, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
      With PicFiles16
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        r = ImageList_Draw(hSIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      With PicFiles32
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        r = ImageList_Draw(hLIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      Set imgObj = ILFiles16(Index).ListImages.Add(, Item.Tag, PicFiles16.Image)
      Set imgObj = ILFiles32(Index).ListImages.Add(, Item.Tag, PicFiles32.Image)
    End If
  Next
  
  Dim strBadIconPath As String
  
  If Right(App.Path, 1) <> "\" Then
    strBadIconPath = App.Path & "\noicon.ico"
  Else
    strBadIconPath = App.Path & "noicon.ico"
  End If
  'we now no where the noiconfile is.
  
  If fso.FileExists(strBadIconPath) = False Then
    Unload Me
  End If
  
      hSIcon = SHGetFileInfo(strBadIconPath, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
      hLIcon = SHGetFileInfo(strBadIconPath, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
      With PicFiles16
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        r = ImageList_Draw(hSIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      With PicFiles32
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        r = ImageList_Draw(hLIcon, SHInfo.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      Set imgObj = ILFiles16(Index).ListImages.Add(, , PicFiles16.Image)
      Set imgObj = ILFiles32(Index).ListImages.Add(, , PicFiles32.Image)
  
  'in the above text we put the badicon.ico in the imagelists
  
  LvFiles(Index).SmallIcons = ILFiles16(Index)
  LvFiles(Index).Icons = ILFiles32(Index)
  'link the imagelists to the filelists
  
  strLastExt = ""
  Dim Image As ListImage
  For Each Item In LvFiles(Index).ListItems
    If fso.FileExists(Item.Tag) Then
      strExt = LCase(fso.GetExtensionName(Item.Tag))
      If strExt <> strLastExt Then
        If Item.Index = 1 Then
          Item.Icon = 1
          Item.SmallIcon = 1
          strLastExt = strExt
        Else
          Item.Icon = LvFiles(Index).ListItems(Item.Index - 1).Icon + 1
          Item.SmallIcon = Item.Icon
          strLastExt = strExt
        End If
      Else
        If strExt <> "exe" And strExt <> "ico" And strExt <> "lnk" Then
          Item.Icon = LvFiles(Index).ListItems(Item.Index - 1).Icon
          Item.SmallIcon = Item.Icon
        Else
          For Each Image In ILFiles16(Index).ListImages
            If Image.Key = "exe" & Item.Index Then
              Item.Icon = Image.Index
              Item.SmallIcon = Image.Index
            End If
          Next
        End If
      End If
    ElseIf fso.FolderExists(Image.Tag) Then
      For Each Image In ILFiles16(Index).ListImages
        If Image.Key = Item.Tag Then
          Item.Icon = Image.Index
          Item.SmallIcon = Image.Index
        End If
      Next
    End If
  Next
  
  'The abovetext pics the right image for each file
  'again using the strLastExt and strExt to make sure
  'the correct image's go to the correct files
  
  For Each Item In LvFiles(Index).ListItems
    If Item.Icon = 0 Then
      Item.Icon = ILFiles32(Index).ListImages.Count
      Item.SmallIcon = ILFiles16(Index).ListImages.Count
    End If
  Next
  
  'For each item without an icon we'll give 'm the noicon
  'icon file, this is just the emptyfile icon included in the
  'directory.
  
  Load_SortList Index, oldSortkey, oldSortorder 'goto phase three, sorting the list
    
  

End Sub

Private Sub Load_SortList(Index As Integer, oldSortkey As Integer, oldSortorder As Integer)
On Error Resume Next
'This procedure is called both when a new folder is explored
'and if a columnheader is clicked
  
  For Each Item In LvFiles(Index).ListItems
    
    If fso.FolderExists(Item.Tag) = True Then
      If Item.Text <> "[..]" Then
        Item.Text = Space(10) & Item.Text
        Item.ListSubItems(1).Text = Item.Text
        Item.ListSubItems(2).Text = Item.Text
        Item.ListSubItems(3).Text = Item.Text
      Else
        If oldSortorder <> lvwAscending Then
          Item.Text = "ZZZ" & Item.Text
          Item.ListSubItems(3).Text = "9999999"
        Else
          Item.Text = Space(10) & Item.Text
          Item.ListSubItems(3).Text = Item.Text
        End If
          Item.ListSubItems(1).Text = Item.Text
          Item.ListSubItems(2).Text = Item.Text
      End If
    ElseIf fso.FileExists(Item.Tag) = True Then
      Set file = fso.GetFile(Item.Tag)
      Item.ListSubItems(2).Text = Left("000000000000000", Len("000000000000000") - Len(file.Size)) & file.Size
      Item.ListSubItems(3).Text = Year(FileDateTime(file.Path)) & Month(FileDateTime(file.Path)) & Day(FileDateTime(file.Path)) & Hour(FileDateTime(file.Path)) & Minute(FileDateTime(file.Path))
    End If
  Next
  
  'the above code makes sure the folders always stick together
  'it also makes sure the "[..]" always stays on top of the list
  
  With LvFiles(Index)
    .SortOrder = oldSortorder
    .SortKey = oldSortkey
    .Sorted = True
  End With
  'this is the default sortingcode
  'no comment actually
  
  
  For Each Item In LvFiles(Index).ListItems
      If Right(Item.Text, 2) = ".]" Then
        Item.Text = "[..]"
        Item.ListSubItems(1).Text = " "
        Item.ListSubItems(2).Text = " "
        Item.ListSubItems(3).Text = " "
      Else
        If fso.FolderExists(Item.Tag) = True Then
          Set Folder = fso.GetFolder(Item.Tag)
          Item.Text = "[" & Folder.Name & "]"
          Item.ListSubItems(1).Text = " "
          Item.ListSubItems(2).Text = " "
          Item.ListSubItems(3).Text = FileDateTime(Folder.Path)
        ElseIf fso.FileExists(Item.Tag) = True Then
          Set file = fso.GetFile(Item.Tag)
          Item.ListSubItems(2).Text = Round(file.Size / 1024) & " kB"
          Item.ListSubItems(3).Text = FileDateTime(file.Path)
        End If
          
      End If
    Next
  'this code looks at every item, decides if there file or folder
  'and then returns the correct info to there textboxes again.
  'this code works fast enough so that users won't see the changing of the text
  
  'I bet there are more ways to sort lists this way, I choose this one
  'if you have a better one yourself... good job!
  'anyway that was the process of loading a folder with icons
  
End Sub
