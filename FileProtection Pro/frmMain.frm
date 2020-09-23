VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Protection Pro"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   11040
      TabIndex        =   21
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Frame Fr 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtHelp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   0
         Width           =   10455
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6135
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   10455
      Begin VB.Image Image1 
         Height          =   2400
         Left            =   0
         Picture         =   "frmMain.frx":164A
         Top             =   0
         Width           =   1980
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Matasurya®"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   480
         Picture         =   "frmMain.frx":31D2
         Top             =   3000
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kutukeyboard@yahoo.com"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblAbout 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   6015
         Left            =   2040
         TabIndex        =   14
         Top             =   120
         Width           =   8415
      End
   End
   Begin VB.Frame Fr 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6135
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton cmd 
         BackColor       =   &H8000000D&
         Caption         =   "Unprotect All"
         Height          =   375
         Index           =   4
         Left            =   8880
         MaskColor       =   &H8000000D&
         TabIndex        =   20
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H8000000D&
         Caption         =   "Unprotect Selected"
         Height          =   375
         Index           =   3
         Left            =   7200
         MaskColor       =   &H8000000D&
         TabIndex        =   12
         Top             =   5760
         Width           =   1575
      End
      Begin VB.ListBox Lst2 
         Height          =   5715
         Left            =   2280
         TabIndex        =   11
         Top             =   0
         Width           =   8175
      End
      Begin VB.ListBox Lst1 
         Height          =   5715
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSComctlLib.ImageList Img 
      Left            =   11880
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   11280
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fr 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   10455
      Begin VB.CommandButton cmd 
         BackColor       =   &H8000000D&
         Caption         =   "Protect All"
         Height          =   375
         Index           =   2
         Left            =   8880
         MaskColor       =   &H8000000D&
         TabIndex        =   19
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H8000000D&
         Caption         =   "Add Folder"
         Height          =   375
         Index           =   1
         Left            =   7320
         MaskColor       =   &H8000000D&
         TabIndex        =   8
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H8000000D&
         Caption         =   "Add Files"
         Height          =   375
         Index           =   0
         Left            =   5760
         MaskColor       =   &H8000000D&
         TabIndex        =   7
         Top             =   5760
         Width           =   1455
      End
      Begin VB.ListBox LstBrowse 
         Height          =   5715
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   10455
      End
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Protected"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Protect More"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   6375
      Left            =   120
      Top             =   480
      Width           =   10695
   End
   Begin VB.Shape Shp 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shp 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shp 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Index           =   2
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shp 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Index           =   3
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f
Dim i As Integer
Dim j As Byte
Private WithEvents iClass As CDrag_Drop
Attribute iClass.VB_VarHelpID = -1
Private size As New CResize

Private Sub iClass_FilesDroped()
Dim v As Integer
With iClass
  For v = 0 To .FileCount - 1
    LstBrowse.AddItem .FileName(v)
  Next
End With
LstBrowse.ListIndex = 0
End Sub

Sub saveLst()
f = FreeFile
If Lst1.ListIndex > 0 Then Lst1.ListIndex = 0
If Lst2.ListIndex > 0 Then Lst2.ListIndex = 0
Open App.Path & "\MTdx.dll" For Output As #f
  For i = 0 To Lst2.ListCount - 1
    Print #f, Lst2.Text
    If Not Lst2.ListIndex = Lst2.ListCount - 1 Then Lst2.ListIndex = Lst2.ListIndex + 1
  Next
Close #f
If Lst1.ListIndex > 0 Then Lst1.ListIndex = 0
If Lst2.ListIndex > 0 Then Lst2.ListIndex = 0
Lst2.SetFocus
End Sub

Private Function BrowseForFolder(ByVal lngHwnd As Long, ByVal strPrompt As String) As String
On Error GoTo ehBrowseForFolder 'Trap for errors
Dim intNull As Integer
Dim lngIDList As Long, lngResult As Long
Dim strPath As String
Dim udtBI As BrowseInfo
With udtBI 'Set API properties (housed in a UDT)
  .lngHwnd = lngHwnd
  .lpszTitle = lstrcat(strPrompt, "")
  .ulFlags = BIF_RETURNONLYFSDIRS
End With
lngIDList = SHBrowseForFolder(udtBI) 'Display the browse folder...
If lngIDList <> 0 Then
  strPath = String(MAX_PATH, 0) 'Create string of nulls so it will fill in with the path\
  lngResult = SHGetPathFromIDList(lngIDList, strPath) 'Retrieves the path selected, places in the null character filled string
  Call CoTaskMemFree(lngIDList) 'Frees memory
  intNull = InStr(strPath, vbNullChar) 'Find the first instance of a null character, so we can get just the path
  If intNull > 0 Then 'Greater than 0 means the path exists...
    strPath = Left(strPath, intNull - 1) 'Set the value
  End If
End If
File1.Path = strPath 'Return the path name
Exit Function 'Abort

ehBrowseForFolder:
BrowseForFolder = Empty 'Return no value
End Function

Private Sub DoEncrypt(sFile As String)
Dim csCrypt       As New clsCrypto
Dim strFile       As String
Dim lFileLength   As Long
lFileLength = FileLen(sFile) ' get length of file to encrypt
strFile = String(lFileLength, vbNullChar) ' allocate string to hold file
Open sFile For Binary Access Read As #1 ' open file in binary
  Get 1, , strFile
Close #1
csCrypt.Password = FrmLog.Tag  ' Get password
csCrypt.InBuffer = strFile
If Not csCrypt.HashFile Then Exit Sub 'generate hash of original file
If Not csCrypt.GeneratePasswordKey Then Exit Sub ' generate password
If Not csCrypt.EncryptFileData Then Exit Sub ' encrypt message data
csCrypt.DestroySessionKey ' destroy key
' check for valid data
If csCrypt.OutBuffer <> "" Then
  Kill sFile ' delete current data file
  Open sFile For Binary Access Write As #2 ' open new file for binary write
    Put 2, , csCrypt.OutBuffer ' write encrypted data to file
  Close #2 ' close open file
End If
End Sub

Private Sub DoDecrypt(sFile As String)
On Error Resume Next
' decrypt file sub
Dim csCrypt     As New clsCrypto
Dim strFile     As String
Dim lFileLength As String
lFileLength = FileLen(sFile) ' get length of file
strFile = String(lFileLength, vbNullChar) ' allocate string to hold file
Open sFile For Binary Access Read As #1 ' open file in binary mode
  Get 1, , strFile
Close #1
csCrypt.Password = FrmLog.Tag  ' set password
csCrypt.InBuffer = strFile
If Not csCrypt.GeneratePasswordKey Then Exit Sub ' generate password
If Not csCrypt.DecryptFileData Then Exit Sub ' decrypt message
csCrypt.DestroySessionKey
' check for valid data
If csCrypt.OutBuffer <> "" Then
Kill sFile ' delete current file
Open sFile For Binary Access Write As #2 ' creat new file
  Put 2, , csCrypt.OutBuffer
Close #2
End If
End Sub

Private Sub cmd_Click(index As Integer)
On Error Resume Next
Select Case index
  Case 0
    Dlg.Filter = "All Files|*.*"
    Dlg.ShowOpen
    txtFileName = Dlg.FileName
    If Dlg.FileName <> "" Then
      LstBrowse.AddItem Dlg.FileName
    End If
  Case 1
    'browse folder
    BrowseForFolder 0, "Browse for folder"
    'If MyPlayer.File1.FileName = "" Then Exit Sub
    File1.ListIndex = 0
    For i = 0 To File1.ListCount - 1
      LstBrowse.AddItem File1.Path & "\" & File1.FileName
      If File1.ListIndex <> File1.ListCount - 1 Then File1.ListIndex = File1.ListIndex + 1
    Next
    LstBrowse.ListIndex = 0
  Case 2
    ' encrypt file
    If LstBrowse.ListIndex < 0 Then
      MsgBox "Please specify file to encrypt!", vbCritical, "No File" ' missing file
      Me.MousePointer = 0
      txtFileName.SetFocus
      Exit Sub
    ElseIf Dir(LstBrowse.Text, vbNormal) = "" Then
      MsgBox "Invalid file name, or file missing!", vbCritical, "Invalid File" ' invalid file or directory
      txtFileName.SetFocus
      Me.MousePointer = 0
      Exit Sub
    Else
      Me.MousePointer = 11
      pbar.Visible = True
      pbar.Value = 0
      pbar.Max = LstBrowse.ListCount
      For i = 0 To LstBrowse.ListCount - 1
        DoEncrypt (LstBrowse.Text)
        If LstBrowse.ListIndex < LstBrowse.ListCount - 1 Then
          LstBrowse.ListIndex = LstBrowse.ListIndex + 1
          pbar.Value = pbar.Value + 1
        End If
      Next
      pbar.Value = 0
      LstBrowse.ListIndex = 0
      For i = 0 To LstBrowse.ListCount - 1
        Lst1.AddItem StripPath(LstBrowse.Text)
        Lst2.AddItem LstBrowse.Text
        If LstBrowse.ListIndex < LstBrowse.ListCount - 1 Then
          LstBrowse.ListIndex = LstBrowse.ListIndex + 1
          pbar.Value = pbar.Value + 1
        End If
      Next
      pbar.Value = 0
      pbar.Visible = False
      Lst1.ListIndex = 0
      saveLst
      MsgBox LstBrowse.ListCount & " files succesfully protected !", vbInformation, "Pregress"
      LstBrowse.Clear
      Me.MousePointer = 0
    End If
  Case 3
    '--decrypt selected
    ' decrypt file
    mx = Lst1.ListIndex
    Me.MousePointer = 11
    DoDecrypt (Lst2.Text)
    Lst1.RemoveItem (mx)
    Lst2.RemoveItem (mx)
    Lst1.ListIndex = 0
    Lst2.ListIndex = 0
    saveLst
    Me.MousePointer = 0
  Case 4
    '-- decript All
     Me.MousePointer = 11
    Lst1.ListIndex = 0
    Lst2.ListIndex = 0
    pbar.Visible = True
    pbar.Value = 0
    pbar.Max = Lst2.ListCount - 1
    For i = 0 To Lst2.ListCount - 1
      DoDecrypt (Lst2.Text)
      If Not Lst2.ListIndex = Lst2.ListCount - 1 Then
        'Lst1.ListIndex = Lst1.ListIndex + 1
        Lst2.ListIndex = Lst2.ListIndex + 1
        pbar.Value = pbar.Value + 1
      End If
    Next
    MsgBox Lst2.ListCount & " files succesfully protected !", vbInformation, "Pregress"
    Lst1.Clear
    Lst2.Clear
    pbar.Value = 0
    pbar.Visible = False
    Kill App.Path & "\MTdx.dll"
    Me.MousePointer = 0
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then ShowPrevInstance
FormTopMost Me.hWnd
Set iClass = New CDrag_Drop
Dim fl As String
f = FreeFile
If FileExists(App.Path & "\MTdx.dll") Then
Open App.Path & "\MTdx.dll" For Input As #f
  Do Until EOF(f)
    Line Input #f, fl
    Lst1.AddItem StripPath(fl)
    Lst2.AddItem fl
  Loop
Close #f
End If
lblAbout = "File Protection Pro is developed by Matasurya, distributed as a freeware so you can feel free to use it." & vbCrLf & _
            "This software has no waranty, so you may use it with your own risk. Just dont edit or move any contents of this software and you'll be fine." & vbCrLf & vbCrLf & _
            "if you forrget your password please contact : kutukeyboard@yahoo.com"
txtHelp = "First add file you want to protect with clicking add files button or add folder buuton " & vbCrLf & _
            "or you can simply drag n drop the file into the listbox at 'Protect More' tab" & vbCrLf & vbCrLf & _
            "Then you can see the protected files list at 'Protected' tab," & vbCrLf & _
            "to unprotec some files, just select the file from the left listbox then click 'unprotect selected' button." & vbCrLf & _
            "To unprotect all files on the list, just simply click 'unprotect all' button." & vbCrLf & vbCrLf & _
            "Thanks for using File Protection Pro." & vbCrLf & _
            "kutukeyboard@yahoo.com with Matasurya software developer"
j = 1
iClass.DragHwnd = LstBrowse.hWnd
iClass.StartDrag
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shp(0).BorderColor = &HFFFFFF
Shp(1).BorderColor = &HFFFFFF
Shp(2).BorderColor = &HFFFFFF
Shp(3).BorderColor = &HFFFFFF
Select Case j
  Case 1
    Shp(0).BackColor = &H80000001
    
    Shp(1).BackColor = &H8000000D
    Shp(2).BackColor = &H8000000D
    Shp(3).BackColor = &H8000000D
  Case 2
    Shp(1).BackColor = &H80000001
    
    Shp(2).BackColor = &H8000000D
    Shp(3).BackColor = &H8000000D
    Shp(0).BackColor = &H8000000D
  Case 3
    Shp(2).BackColor = &H80000001
    
    Shp(3).BackColor = &H8000000D
    Shp(0).BackColor = &H8000000D
    Shp(1).BackColor = &H8000000D
  Case 4
    Shp(3).BackColor = &H80000001
    
    Shp(0).BackColor = &H8000000D
    Shp(1).BackColor = &H8000000D
    Shp(2).BackColor = &H8000000D
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 0
iClass.StopDrag
Unload FrmLog
End
Cancel = 1
End Sub

Private Sub lbl_Click(index As Integer)
Select Case index
  Case 0
    Fr(0).Visible = True
    Fr(1).Visible = False
    Fr(2).Visible = False
    Fr(3).Visible = False
    j = 1
  Case 1
    Fr(1).Visible = True
    Fr(2).Visible = False
    Fr(3).Visible = False
    Fr(0).Visible = False
    j = 2
  Case 2
    Fr(2).Visible = True
    Fr(3).Visible = False
    Fr(0).Visible = False
    Fr(1).Visible = False
    j = 3
  Case 3
    Fr(3).Visible = True
    Fr(0).Visible = False
    Fr(1).Visible = False
    Fr(2).Visible = False
    j = 4
End Select
End Sub

Private Sub lbl_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Shp(lbl(index).index).BackColor = &H80000002
Shp(lbl(index).index).BorderColor = &H80000003
End Sub

Private Sub lst1_Click()
Lst2.ListIndex = Lst1.ListIndex
End Sub

Private Sub Lst1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lst2.ListIndex = Lst1.ListIndex
End Sub

Private Sub LstBrowse_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 46 Then
  LstBrowse.RemoveItem (LstBrowse.ListIndex)
End If
If LstBrowse.ListCount <> 0 Then
  LstBrowse.ListIndex = 0
End If
End Sub
