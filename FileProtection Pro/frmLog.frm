VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLog 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Who's The Hell Are You?"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3645
   ForeColor       =   &H8000000F&
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer tmr 
      Left            =   2160
      Top             =   2400
   End
   Begin VB.TextBox txt1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "."
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Magic Word Please"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FrmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f
Dim mPass As String
Private Declare Function PaintDesktop Lib "User32" _
(ByVal hdc As Long) As Long

Private Sub DoEncrypt(sFile As String)
Dim csCrypt       As New clsCrypto
Dim strFile       As String
Dim lFileLength   As Long
lFileLength = FileLen(sFile) ' get length of file to encrypt
strFile = String(lFileLength, vbNullChar) ' allocate string to hold file
Open sFile For Binary Access Read As #1 ' open file in binary
  Get 1, , strFile
Close #1
csCrypt.Password = "matasurya"  ' Get password
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
csCrypt.Password = "matasurya"  ' set password
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

Private Sub Form_Load()
On Error Resume Next
App.TaskVisible = False
If FileExists(App.Path & "\Mse32.dll") Then
  DoDecrypt (App.Path & "\Mse32.dll")
  f = FreeFile
  Open App.Path & "\Mse32.dll" For Input As #f
    Line Input #f, mPass
  Close #f
  Me.Tag = mPass
  DoEncrypt (App.Path & "\Mse32.dll")
Else
  Me.Visible = False
  mPass = InputBox("Please Create your password here !", "Create password")
  If mPass = "" Then
      MsgBox "You must add a password first !" & vbCrLf & _
            "Sorry you must reload this software !", vbExclamation, "Password Required"
      End
  End If
  f = FreeFile
  Open App.Path & "\Mse32.dll" For Output As #f
    Print #f, mPass
  Close #f
  Me.Tag = mPass
  DoEncrypt (App.Path & "\Mse32.dll")
  frmMain.Show
  Me.Hide
End If
End Sub

Private Sub txt1_Change()
If txt1 = Me.Tag Then
  frmMain.Show
  Me.Hide
End If
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
If keysacii = vbKeyReturn Then
  If txt1 = Me.Tag Then
  frmMain.Show
Me.Hide
End If
End If
End Sub
